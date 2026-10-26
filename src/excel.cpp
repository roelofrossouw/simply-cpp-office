#include "excel.h"
#include "excel_sheet_names.h"

#include <sc_archive.h>
#include <xml.h>

#include <atomic>
#include <cctype>
#include <charconv>
#include <cmath>
#include <cstdio>
#include <fstream>
#include <iterator>
#include <locale>
#include <map>
#include <random>
#include <sstream>
#include <stdexcept>
#include <utility>
#include <vector>

namespace sc::office::excel {
    namespace {
        // An attribute by its local name, whatever its namespace prefix ("r:id" is usual).
        std::string attribute_named(const web::element &element, const std::string_view local) {
            for (const auto &each: element.attributes()) {
                const std::string_view name = each.name;
                const auto colon = name.find(':');
                if ((colon == std::string_view::npos ? name : name.substr(colon + 1)) == local) return std::string{each.value};
            }
            return {};
        }

        bool is_hex(const char c) { return std::isxdigit(static_cast<unsigned char>(c)) != 0; }

        // Whether text holds "_xHHHH_" at at: how Excel writes a character XML can't hold.
        bool escape_at(const std::string_view text, const std::size_t at) {
            return text.size() >= at + 7 && text.compare(at, 2, "_x") == 0 && is_hex(text[at + 2]) && is_hex(text[at + 3]) &&
                   is_hex(text[at + 4]) && is_hex(text[at + 5]) && text[at + 6] == '_';
        }

        void append_utf8(std::string &out, const unsigned code) {
            if (code < 0x80) {
                out += static_cast<char>(code);
            } else if (code < 0x800) {
                out += static_cast<char>(0xC0 | (code >> 6));
                out += static_cast<char>(0x80 | (code & 0x3F));
            } else {
                out += static_cast<char>(0xE0 | (code >> 12));
                out += static_cast<char>(0x80 | ((code >> 6) & 0x3F));
                out += static_cast<char>(0x80 | (code & 0x3F));
            }
        }

        // Text as Excel stored it, with its "_xHHHH_" escapes undone.
        std::string unescaped(const std::string_view text) {
            std::string out;
            out.reserve(text.size());
            for (std::size_t at = 0; at < text.size();) {
                if (escape_at(text, at)) {
                    append_utf8(out, static_cast<unsigned>(std::stoul(std::string{text.substr(at + 2, 4)}, nullptr, 16)));
                    at += 7;
                } else {
                    out += text[at++];
                }
            }
            return out;
        }

        // A shared or inline string (<si>, <is>): its <t>, or its runs' <t>s; phonetic hints left out.
        std::string string_item(const web::element &item) {
            std::string text;
            for (const auto &part: item.children()) {
                if (part.is("t")) {
                    text += part.text();
                } else if (part.is("r")) {
                    if (const auto t = part.child("t")) text += t->text();
                }
            }
            return unescaped(text);
        }

        double number(const std::string_view text) {
            std::istringstream in{std::string{text}};
            in.imbue(std::locale::classic());
            double result = 0;
            in >> result;
            if (in.fail()) throw std::runtime_error("excel: \"" + std::string{text} + "\" isn't a number");
            return result;
        }

        // A part's path from a relationship's target, relative to the folder of the part it's from.
        std::string part_path(const std::string_view folder, const std::string_view target) {
            if (target.starts_with('/')) return std::string{target.substr(1)};
            std::string path{folder};
            std::string_view rest = target;
            while (rest.starts_with("../")) {
                rest.remove_prefix(3);
                if (const auto slash = path.find_last_of('/', path.size() - 2); slash != std::string::npos) path.erase(slash + 1);
                else path.clear();
            }
            return path + std::string{rest};
        }

        // A formula moved by rows and columns, as Excel moves one it keeps once for a range (a shared
        // formula): its relative references move, $-fixed parts and text in quotes don't.
        std::string moved(const std::string_view formula, const long rows, const long columns) {
            const auto word = [](const char c) { return std::isalnum(static_cast<unsigned char>(c)) || c == '_' || c == '.'; };
            std::string out;
            for (std::size_t at = 0; at < formula.size();) {
                const char c = formula[at];
                if (c == '"' || c == '\'') { // text, or a sheet's name: as it is, its doubled quotes too
                    auto end = at + 1;
                    while (end < formula.size() && (formula[end] != c || (end + 1 < formula.size() && formula[end + 1] == c)))
                        end += formula[end] == c ? 2 : 1;
                    out += formula.substr(at, end + 1 - at);
                    at = end + 1;
                    continue;
                }
                if (!std::isalpha(static_cast<unsigned char>(c)) && c != '$') {
                    out += c;
                    ++at;
                    continue;
                }
                // A reference: [$]letters[$]digits, not part of a longer word or a function's name.
                auto end = at;
                const bool column_fixed = formula[end] == '$';
                if (column_fixed) ++end;
                const auto letters_at = end;
                while (end < formula.size() && std::isalpha(static_cast<unsigned char>(formula[end]))) ++end;
                const auto letters = formula.substr(letters_at, end - letters_at);
                const bool row_fixed = end < formula.size() && formula[end] == '$';
                if (row_fixed) ++end;
                const auto digits_at = end;
                while (end < formula.size() && std::isdigit(static_cast<unsigned char>(formula[end]))) ++end;
                const auto digits = formula.substr(digits_at, end - digits_at);
                const bool after_word = at > 0 && word(formula[at - 1]); // the E5 in 1E5
                const bool reference = !after_word && !letters.empty() && letters.size() <= 3 && !digits.empty() &&
                                       (end == formula.size() || (!word(formula[end]) && formula[end] != '(' && formula[end] != '!'));
                if (!reference) { // a word: a function, a name, TRUE ...
                    auto stop = at + 1;
                    while (stop < formula.size() && (word(formula[stop]) || formula[stop] == '$')) ++stop;
                    out += formula.substr(at, stop - at);
                    at = stop;
                    continue;
                }
                const address was{std::string{letters} + std::string{digits}};
                const long column = column_fixed ? static_cast<long>(was.column) : static_cast<long>(was.column) + columns;
                const long row = row_fixed ? static_cast<long>(was.row) : static_cast<long>(was.row) + rows;
                if (column < 0 || row < 0 || column >= static_cast<long>(max_columns) || row >= static_cast<long>(max_rows)) {
                    out += "#REF!";
                } else {
                    out += (column_fixed ? "$" : "") + column_name(static_cast<std::size_t>(column)) + (row_fixed ? "$" : "") +
                           std::to_string(row + 1);
                }
                at = end;
            }
            return out;
        }

        // The formulas a sheet keeps once for a range, by their si: the first cell's formula and place.
        struct shared_formula {
            std::string text;
            address from;
        };

        value cell_value(const web::element &c, const std::vector<std::string> &shared) {
            const std::string_view type = c["t"];
            if (type == "inlineStr") {
                if (const auto is = c.child("is")) return string_item(*is);
                return {};
            }
            const auto v = c.child("v");
            if (!v) return {};
            const auto text = v->text();
            if (type == "s") {
                const auto index = static_cast<std::size_t>(number(text));
                if (index >= shared.size()) throw std::runtime_error("excel: shared string " + text + " isn't there");
                return shared[index];
            }
            if (type == "b") return text == "1" || text == "true";
            if (type == "str" || type == "e") return unescaped(text);
            if (text.empty()) return {};
            return number(text);
        }

        // A cell's formula, with "=": its own, or the one it shares, moved to it.
        std::string cell_formula(const web::element &c, const address where, std::map<std::string, shared_formula> &formulas) {
            const auto f = c.child("f");
            if (!f) return {};
            auto text = unescaped(f->text());
            if ((*f)["t"] == "shared") {
                const std::string si{(*f)["si"]};
                if (!text.empty()) {
                    formulas[si] = {text, where};
                } else if (const auto found = formulas.find(si); found != formulas.end()) {
                    text = moved(found->second.text, static_cast<long>(where.row) - static_cast<long>(found->second.from.row),
                                 static_cast<long>(where.column) - static_cast<long>(found->second.from.column));
                }
            }
            return text.empty() ? std::string{} : "=" + text;
        }

        sheet sheet_of(std::string name, const archive::reader &book, const std::string &path, const std::vector<std::string> &shared) {
            sheet out{std::move(name)};
            if (!book.contains(path)) return out;
            const auto document = web::document::parse(book.read(path));
            const auto data = document.root().child("sheetData");
            if (!data) return out;
            std::map<std::string, shared_formula> formulas;
            std::size_t next_row = 0;
            for (const auto &row: data->children("row")) {
                std::size_t row_index = next_row;
                if (const auto r = row["r"]; !r.empty()) row_index = static_cast<std::size_t>(number(r)) - 1;
                next_row = row_index + 1;
                std::size_t next_column = 0;
                for (const auto &c: row.children("c")) {
                    address where{row_index, next_column};
                    if (const auto r = c["r"]; !r.empty()) where = address{r};
                    next_column = where.column + 1;
                    auto what = cell_value(c, shared);
                    auto formula = cell_formula(c, where, formulas);
                    if (std::holds_alternative<std::monostate>(what) && formula.empty()) continue;
                    auto &cell = out[where];
                    cell.value = std::move(what);
                    cell.formula = std::move(formula);
                }
            }
            return out;
        }

        std::deque<sheet> sheets_of(const archive::reader &book) {
            if (!book.contains("xl/workbook.xml")) throw std::runtime_error("excel: not an .xlsx (no xl/workbook.xml)");
            std::map<std::string, std::string> targets; // relationship id -> part
            if (book.contains("xl/_rels/workbook.xml.rels")) {
                const auto rels = web::document::parse(book.read("xl/_rels/workbook.xml.rels"));
                for (const auto &relationship: rels.root().children("Relationship"))
                    if (relationship["TargetMode"] != "External")
                        targets[std::string{relationship["Id"]}] = part_path("xl/", relationship["Target"]);
            }
            std::vector<std::string> shared;
            if (book.contains("xl/sharedStrings.xml")) {
                const auto strings = web::document::parse(book.read("xl/sharedStrings.xml"));
                for (const auto &item: strings.root().children("si")) shared.push_back(string_item(item));
            }

            std::deque<sheet> out;
            const auto document = web::document::parse(book.read("xl/workbook.xml"));
            if (const auto sheets = document.root().child("sheets")) {
                for (const auto &each: sheets->children("sheet")) {
                    const auto target = targets.find(attribute_named(each, "id"));
                    out.push_back(sheet_of(std::string{each["name"]}, book, target == targets.end() ? std::string{} : target->second, shared));
                }
            }
            return out;
        }

        // Text for XML, with the characters XML can't hold written the way Excel does ("_x0001_"),
        // and "_xHHHH_" already in the text kept as it is.
        std::string xml_text(const std::string_view text) {
            std::string out;
            out.reserve(text.size());
            for (std::size_t at = 0; at < text.size(); ++at) {
                const auto c = static_cast<unsigned char>(text[at]);
                if (c == '&') out += "&amp;";
                else if (c == '<') out += "&lt;";
                else if (c == '>') out += "&gt;";
                else if (c == '"') out += "&quot;";
                else if (c == '_' && escape_at(text, at)) out += "_x005F_";
                else if (c < 0x20 && c != '\t' && c != '\n' && c != '\r') {
                    char hex[8];
                    std::snprintf(hex, sizeof hex, "_x%04X_", c);
                    out += hex;
                } else out += static_cast<char>(c);
            }
            return out;
        }

        std::string number_text(const double number) {
            if (!std::isfinite(number)) throw std::invalid_argument("excel: a cell's number must be finite");
            char text[32];
            const auto [end, error] = std::to_chars(text, text + sizeof text, number);
            return {text, end};
        }

        constexpr std::string_view xml_declaration = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)" "\n";
        constexpr std::string_view relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        // A cell's XML: its formula (without the =) and the value it holds.
        std::string cell_xml(const cell &each) {
            std::string formula;
            if (!each.formula.empty())
                formula = "<f>" + xml_text(each.formula.starts_with('=') ? std::string_view{each.formula}.substr(1) : std::string_view{each.formula}) + "</f>";
            const auto start = R"(<c r=")" + each.name() + '"';
            if (const auto *text = std::get_if<std::string>(&each.value)) {
                if (!formula.empty()) return start + R"( t="str">)" + formula + "<v>" + xml_text(*text) + "</v></c>";
                return start + R"( t="inlineStr"><is><t xml:space="preserve">)" + xml_text(*text) + "</t></is></c>";
            }
            if (const auto *truth = std::get_if<bool>(&each.value)) return start + R"( t="b">)" + formula + "<v>" + (*truth ? "1" : "0") + "</v></c>";
            if (const auto *number = std::get_if<double>(&each.value)) return start + ">" + formula + "<v>" + number_text(*number) + "</v></c>";
            return start + ">" + formula + "</c>";
        }

        std::string sheet_xml(const sheet &each) {
            std::string xml{xml_declaration};
            xml += R"(<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>)";
            for (const auto &row: each) {
                std::string cells;
                for (const auto &cell: row)
                    if (!cell.empty()) cells += cell_xml(cell);
                if (!cells.empty()) xml += R"(<row r=")" + std::to_string(row.index() + 1) + R"(">)" + cells + "</row>";
            }
            return xml + "</sheetData></worksheet>";
        }
    }

    workbook read(const std::filesystem::path &file) {
        workbook book;
        book.sheets_ = sheets_of(archive::open(file));
        return book;
    }

    workbook read_from(std::string bytes) {
        workbook book;
        book.sheets_ = sheets_of(archive::from_memory(std::move(bytes)));
        return book;
    }

    void write(const workbook &book, const std::filesystem::path &file) {
        if (book.empty()) throw std::invalid_argument("excel: a workbook needs a sheet");
        std::vector<std::string_view> names;
        bool formulas = false;
        for (const auto &each: book) {
            detail::check_name(each.name());
            for (const auto name: names)
                if (detail::same_name(each.name(), name)) throw std::invalid_argument("excel: two sheets are named \"" + each.name() + "\"");
            names.push_back(each.name());
            for (const auto &row: each)
                for (const auto &cell: row) formulas = formulas || !cell.formula.empty();
        }

        // Everything is made first, so a workbook that can't be written leaves no file behind.
        std::vector<std::pair<std::string, std::string>> parts;
        std::string types{xml_declaration};
        types += R"(<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">)"
                 R"(<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>)"
                 R"(<Default Extension="xml" ContentType="application/xml"/>)"
                 R"(<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>)";
        std::string sheets, rels{xml_declaration};
        rels += R"(<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">)";
        for (std::size_t i = 1; i <= book.size(); ++i) {
            const auto n = std::to_string(i);
            types += R"(<Override PartName="/xl/worksheets/sheet)" + n +
                     R"(.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>)";
            sheets += R"(<sheet name=")" + xml_text(names[i - 1]) + R"(" sheetId=")" + n + R"(" r:id="rId)" + n + R"("/>)";
            rels += R"(<Relationship Id="rId)" + n + R"(" Type=")" + std::string{relationships} + R"(/worksheet" Target="worksheets/sheet)" + n + R"(.xml"/>)";
        }
        parts.emplace_back("[Content_Types].xml", types + "</Types>");
        parts.emplace_back("_rels/.rels", std::string{xml_declaration} +
                                              R"(<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">)"
                                              R"(<Relationship Id="rId1" Type=")" + std::string{relationships} +
                                              R"(/officeDocument" Target="xl/workbook.xml"/></Relationships>)");
        parts.emplace_back("xl/workbook.xml", std::string{xml_declaration} +
                                                  R"(<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r=")" +
                                                  std::string{relationships} + R"("><sheets>)" + sheets + "</sheets>" +
                                                  (formulas ? R"(<calcPr fullCalcOnLoad="1"/>)" : "") + "</workbook>");
        parts.emplace_back("xl/_rels/workbook.xml.rels", rels + "</Relationships>");
        std::size_t n = 0;
        for (const auto &each: book) parts.emplace_back("xl/worksheets/sheet" + std::to_string(++n) + ".xml", sheet_xml(each));

        auto out = archive::create(file);
        for (auto &[name, bytes]: parts) out.add(name, std::move(bytes));
        out.close();
    }

    std::string bytes(const workbook &book) {
        // sc-archive writes ZIP archives to files: one of our own, read back and removed.
        static std::atomic<unsigned> made{0};
        const auto file = std::filesystem::temp_directory_path() /
                          ("sc-office-excel-" + std::to_string(std::random_device{}()) + "-" + std::to_string(made++) + ".xlsx");
        struct removed {
            std::filesystem::path file;
            ~removed() {
                std::error_code ignored;
                std::filesystem::remove(file, ignored);
            }
        } cleanup{file};
        write(book, file);
        std::ifstream in{file, std::ios::binary};
        return {std::istreambuf_iterator<char>{in}, {}};
    }
}
