#include "excel.h"
#include "excel_sheet_names.h"

#include <sc_archive.h>
#include <xml.h>

#include <algorithm>
#include <atomic>
#include <cctype>
#include <charconv>
#include <cmath>
#include <cstdio>
#include <fstream>
#include <iterator>
#include <locale>
#include <map>
#include <optional>
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

        // Excel's 64 indexed colours, for files that name colours by number (indexed="10").
        constexpr unsigned indexed_colors[64] = {
            0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, 0x000000, 0xFFFFFF, 0xFF0000,
            0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, 0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080,
            0xC0C0C0, 0x808080, 0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF, 0x000080,
            0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF, 0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99,
            0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99, 0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699,
            0x969696, 0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333};

        sc::color rgb_color(const unsigned rgb) { return sc::color::from_web(static_cast<int>(rgb >> 16), static_cast<int>((rgb >> 8) & 0xFF), static_cast<int>(rgb & 0xFF)); }

        // "FFC00000" (alpha first) or "C00000".
        std::optional<sc::color> hex_color(const std::string_view text) {
            const auto hex = text.size() == 8 ? text.substr(2) : text;
            unsigned rgb = 0;
            if (hex.size() != 6 || std::from_chars(hex.data(), hex.data() + 6, rgb, 16).ptr != hex.data() + 6) return std::nullopt;
            return rgb_color(rgb);
        }

        // A colour lightened (tint above 0) or darkened (below), as Excel does: its HSL lightness moved
        // toward white or black.
        sc::color tinted(const sc::color &color, const double tint) {
            const double r = color.red(), g = color.green(), b = color.blue();
            const double high = std::max({r, g, b}), low = std::min({r, g, b});
            double hue = 0, saturation = 0, lightness = (high + low) / 2;
            if (high != low) {
                const double range = high - low;
                saturation = lightness > 0.5 ? range / (2 - high - low) : range / (high + low);
                hue = high == r ? (g - b) / range + (g < b ? 6 : 0) : high == g ? (b - r) / range + 2 : (r - g) / range + 4;
                hue /= 6;
            }
            lightness = tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint;
            if (saturation == 0) return {lightness, lightness, lightness};
            const double q = lightness < 0.5 ? lightness * (1 + saturation) : lightness + saturation - lightness * saturation;
            const double p = 2 * lightness - q;
            const auto channel = [&](double t) {
                if (t < 0) t += 1;
                if (t > 1) t -= 1;
                if (t < 1.0 / 6) return p + (q - p) * 6 * t;
                if (t < 0.5) return q;
                if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
                return p;
            };
            return {channel(hue + 1.0 / 3), channel(hue), channel(hue - 1.0 / 3)};
        }

        // A colour as styles.xml gives one: rgb="FFC00000", theme="4" (with a tint="0.4"), indexed="10".
        // Automatic, or one that can't be told, is none.
        std::optional<sc::color> color_of(const web::element &c, const std::vector<sc::color> &theme) {
            std::optional<sc::color> found;
            if (const auto rgb = c["rgb"]; !rgb.empty()) {
                found = hex_color(rgb);
            } else if (const auto index = c["theme"]; !index.empty()) {
                if (const auto at = static_cast<std::size_t>(number(index)); at < theme.size()) found = theme[at];
            } else if (const auto index = c["indexed"]; !index.empty()) {
                if (const auto at = static_cast<std::size_t>(number(index)); at < std::size(indexed_colors)) found = rgb_color(indexed_colors[at]);
            }
            if (const auto tint = c["tint"]; found && !tint.empty() && number(tint) != 0) found = tinted(*found, std::clamp(number(tint), -1.0, 1.0));
            return found;
        }

        // The theme's colours in the order styles.xml counts them: light 1, dark 1, light 2, dark 2,
        // the six accents, then the two hyperlink colours.
        std::vector<sc::color> theme_colors(const archive::reader &book, const std::string &path) {
            if (path.empty() || !book.contains(path)) return {};
            const auto document = web::document::parse(book.read(path));
            const auto elements = document.root().child("themeElements");
            const auto scheme = elements ? elements->child("clrScheme") : std::nullopt;
            if (!scheme) return {};
            std::map<std::string, sc::color, std::less<>> named;
            for (const auto &each: scheme->children()) {
                std::optional<sc::color> found;
                if (const auto rgb = each.child("srgbClr")) found = hex_color((*rgb)["val"]);
                else if (const auto system = each.child("sysClr")) found = hex_color((*system)["lastClr"]);
                if (found) named.emplace(std::string{each.name()}, *found);
            }
            std::vector<sc::color> out;
            for (const auto name: {"lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"}) {
                const auto found = named.find(name);
                if (found == named.end()) break;
                out.push_back(found->second);
            }
            return out;
        }

        template<typename Kind>
        Kind named_or(Kind (*named)(std::string_view), const std::string_view name, const Kind otherwise) {
            try {
                return named(name);
            } catch (const std::invalid_argument &) {
                return otherwise;
            }
        }

        // The number formats Excel has without saying (numFmtId below 164) that read the same
        // everywhere; 14 is the reader's own short date ("mm-dd-yy" stands for it), so it's only read.
        constexpr std::pair<unsigned, std::string_view> builtin_formats[] = {
            {1, "0"}, {2, "0.00"}, {3, "#,##0"}, {4, "#,##0.00"}, {9, "0%"}, {10, "0.00%"}, {11, "0.00E+00"}, {12, "# ?/?"},
            {13, "# ?\?/??"}, {14, "mm-dd-yy"}, {15, "d-mmm-yy"}, {16, "d-mmm"}, {17, "mmm-yy"}, {18, "h:mm AM/PM"},
            {19, "h:mm:ss AM/PM"}, {20, "h:mm"}, {21, "h:mm:ss"}, {22, "m/d/yy h:mm"}, {37, "#,##0 ;(#,##0)"},
            {38, "#,##0 ;[Red](#,##0)"}, {39, "#,##0.00;(#,##0.00)"}, {40, "#,##0.00;[Red](#,##0.00)"}, {45, "mm:ss"},
            {46, "[h]:mm:ss"}, {47, "mmss.0"}, {48, "##0.0E+0"}, {49, "@"}};

        // What Excel names so, or otherwise when a file has a name it doesn't have (it's drawn anyway).
        template<typename Kind>
        Kind named_or(Kind (*named)(std::string_view), const std::string_view name, const Kind otherwise);

        // An xf's <alignment horizontal="center" vertical="top" wrapText="1" indent="2"/>.
        alignment alignment_of(const web::element &xf) {
            alignment made;
            const auto element = xf.child("alignment");
            if (!element) return made;
            if (const auto kind = (*element)["horizontal"]; !kind.empty()) made.horizontal = named_or(horizontal_named, kind, horizontal::general);
            if (const auto kind = (*element)["vertical"]; !kind.empty()) made.vertical = named_or(vertical_named, kind, vertical::bottom);
            made.wrap = (*element)["wrapText"] == "1" || (*element)["wrapText"] == "true";
            if (const auto indent = (*element)["indent"]; !indent.empty()) made.indent = static_cast<unsigned>(std::clamp(number(indent), 0.0, 250.0));
            // textRotation: 0 to 90 up, 91 to 180 down (90 less it), 255 stacked.
            if (const auto rotation = (*element)["textRotation"]; !rotation.empty()) {
                const auto degrees = static_cast<int>(number(rotation));
                if (degrees == 255) made.stacked = true;
                else if (degrees > 90 && degrees <= 180) made.rotation = 90 - degrees;
                else if (degrees >= 0 && degrees <= 90) made.rotation = degrees;
            }
            return made;
        }

        // Whether a font's <b/>, <i/> ... is there and on (val="0" or "false" turns it off).
        bool on(const std::optional<web::element> &flag) { return flag && (*flag)["val"] != "0" && (*flag)["val"] != "false"; }

        // A border's side: <left style="thin"><color rgb="FF000000"/></left>.
        edge edge_of(const std::optional<web::element> &side, const std::vector<sc::color> &theme) {
            edge made;
            if (!side || (*side)["style"].empty()) return made;
            made.line = named_or(line_named, (*side)["style"], line::thin);
            if (const auto color = side->child("color")) made.color = color_of(*color, theme);
            return made;
        }

        // The styles cells point to (s="3": cellXfs' fourth), as excel::style: what's the same as the
        // workbook's default font is left unset.
        std::vector<style> styles_of(const archive::reader &book, const std::string &path, const std::vector<sc::color> &theme) {
            if (path.empty() || !book.contains(path)) return {};
            const auto document = web::document::parse(book.read(path));
            const auto &root = document.root();
            std::vector<font> fonts;
            if (const auto list = root.child("fonts")) {
                for (const auto &each: list->children("font")) {
                    font made;
                    if (const auto size = each.child("sz"); size && !(*size)["val"].empty()) made.size = number((*size)["val"]);
                    if (const auto color = each.child("color")) made.color = color_of(*color, theme);
                    made.bold = on(each.child("b"));
                    made.italic = on(each.child("i"));
                    made.strike = on(each.child("strike"));
                    if (const auto underline = each.child("u")) {
                        const auto kind = (*underline)["val"];
                        made.underline = kind.empty() ? underline::single : named_or(underline_named, kind, underline::single);
                    }
                    fonts.push_back(made);
                }
            }
            if (!fonts.empty()) {
                const auto plain = fonts.front();
                for (auto &each: fonts) {
                    if (each.size && plain.size && each.size->points() == plain.size->points()) each.size.reset();
                    if (each.color == plain.color) each.color.reset();
                }
            }
            std::map<unsigned, std::string> formats{builtin_formats, builtin_formats + std::size(builtin_formats)};
            if (const auto list = root.child("numFmts"))
                for (const auto &each: list->children("numFmt"))
                    if (!each["numFmtId"].empty()) formats[static_cast<unsigned>(number(each["numFmtId"]))] = unescaped(each["formatCode"]);
            std::vector<excel::borders> borders;
            if (const auto list = root.child("borders")) {
                for (const auto &each: list->children("border")) {
                    // start and end are what newer files call left and right.
                    const auto left = each.child("left") ? each.child("left") : each.child("start");
                    const auto right = each.child("right") ? each.child("right") : each.child("end");
                    const auto diagonal = edge_of(each.child("diagonal"), theme);
                    const auto yes = [&](const std::string_view name) { return each[name] == "1" || each[name] == "true"; };
                    borders.push_back({edge_of(left, theme), edge_of(right, theme), edge_of(each.child("top"), theme), edge_of(each.child("bottom"), theme),
                                       yes("diagonalUp") ? diagonal : edge{}, yes("diagonalDown") ? diagonal : edge{}});
                }
            }
            std::vector<std::optional<sc::color>> fills;
            if (const auto list = root.child("fills")) {
                for (const auto &each: list->children("fill")) {
                    std::optional<sc::color> background;
                    if (const auto pattern = each.child("patternFill"); pattern && (*pattern)["patternType"] == "solid")
                        if (const auto color = pattern->child("fgColor")) background = color_of(*color, theme);
                    fills.push_back(background);
                }
            }
            std::vector<style> out;
            if (const auto list = root.child("cellXfs")) {
                for (const auto &xf: list->children("xf")) {
                    style made;
                    if (const auto id = xf["fontId"]; !id.empty())
                        if (const auto at = static_cast<std::size_t>(number(id)); at < fonts.size()) made.font = fonts[at];
                    if (const auto id = xf["numFmtId"]; !id.empty())
                        if (const auto found = formats.find(static_cast<unsigned>(number(id))); found != formats.end() && found->second != "General")
                            made.number_format = found->second;
                    made.alignment = alignment_of(xf);
                    if (const auto id = xf["borderId"]; !id.empty())
                        if (const auto at = static_cast<std::size_t>(number(id)); at < borders.size()) made.borders = borders[at];
                    if (const auto id = xf["fillId"]; !id.empty())
                        if (const auto at = static_cast<std::size_t>(number(id)); at < fills.size()) made.background = fills[at];
                    out.push_back(made);
                }
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

        sheet sheet_of(std::string name, const archive::reader &book, const std::string &path, const std::vector<std::string> &shared,
                       const std::vector<style> &styles) {
            sheet out{std::move(name)};
            if (!book.contains(path)) return out;
            const auto document = web::document::parse(book.read(path));
            // Columns given a width; others (styled only) keep Excel's default.
            if (const auto columns = document.root().child("cols")) {
                for (const auto &col: columns->children("col")) {
                    const auto yes = [&](const std::string_view name) { return col[name] == "1" || col[name] == "true"; };
                    const bool sized = !col["width"].empty() && yes("customWidth"), hidden = yes("hidden");
                    if ((!sized && !hidden) || col["min"].empty() || col["max"].empty()) continue;
                    const auto width = sized ? std::clamp(number(col["width"]), 0.0, 255.0) : 0.0;
                    const auto first = static_cast<std::size_t>(number(col["min"]));
                    const auto last = std::min(static_cast<std::size_t>(number(col["max"])), max_columns);
                    for (auto column = std::max<std::size_t>(first, 1); column <= last; ++column) {
                        if (sized) out.width(column - 1, width);
                        if (hidden) out.hidden_column(column - 1, true);
                    }
                }
            }
            // Frozen panes: so many rows (ySplit) and columns (xSplit) kept in view.
            if (const auto views = document.root().child("sheetViews"))
                if (const auto view = views->child("sheetView"))
                    if (const auto pane = view->child("pane"); pane && ((*pane)["state"] == "frozen" || (*pane)["state"] == "frozenSplit")) {
                        const auto split = [&](const std::string_view name) {
                            return (*pane)[name].empty() ? std::size_t{0} : static_cast<std::size_t>(std::max(number((*pane)[name]), 0.0));
                        };
                        if (split("ySplit") < max_rows && split("xSplit") < max_columns) out.freeze({split("ySplit"), split("xSplit")});
                    }
            // Merged cells; one that isn't a range, or overlaps one before it, is left out.
            if (const auto merges = document.root().child("mergeCells")) {
                for (const auto &each: merges->children("mergeCell")) {
                    try {
                        out.merge(range{each["ref"]});
                    } catch (const std::exception &) {
                    }
                }
            }
            const auto data = document.root().child("sheetData");
            if (!data) return out;
            std::map<std::string, shared_formula> formulas;
            std::size_t next_row = 0;
            for (const auto &row: data->children("row")) {
                std::size_t row_index = next_row;
                if (const auto r = row["r"]; !r.empty()) row_index = static_cast<std::size_t>(number(r)) - 1;
                next_row = row_index + 1;
                if (const auto custom = row["customHeight"]; (custom == "1" || custom == "true") && !row["ht"].empty() && row_index < max_rows)
                    out.height(row_index, std::clamp(number(row["ht"]), 0.0, 409.0));
                if ((row["hidden"] == "1" || row["hidden"] == "true") && row_index < max_rows) out.hidden_row(row_index, true);
                std::size_t next_column = 0;
                for (const auto &c: row.children("c")) {
                    address where{row_index, next_column};
                    if (const auto r = c["r"]; !r.empty()) where = address{r};
                    next_column = where.column + 1;
                    auto what = cell_value(c, shared);
                    auto formula = cell_formula(c, where, formulas);
                    style look;
                    if (const auto s = c["s"]; !s.empty())
                        if (const auto at = static_cast<std::size_t>(number(s)); at < styles.size()) look = styles[at];
                    if (std::holds_alternative<std::monostate>(what) && formula.empty() && look.empty()) continue;
                    auto &cell = out[where];
                    cell.value = std::move(what);
                    cell.formula = std::move(formula);
                    cell.style = std::move(look);
                }
            }
            return out;
        }

        std::deque<sheet> sheets_of(const archive::reader &book) {
            if (!book.contains("xl/workbook.xml")) throw std::runtime_error("excel: not an .xlsx (no xl/workbook.xml)");
            std::map<std::string, std::string> targets; // relationship id -> part
            std::string styles_path, theme_path;
            if (book.contains("xl/_rels/workbook.xml.rels")) {
                const auto rels = web::document::parse(book.read("xl/_rels/workbook.xml.rels"));
                for (const auto &relationship: rels.root().children("Relationship")) {
                    if (relationship["TargetMode"] == "External") continue;
                    auto path = part_path("xl/", relationship["Target"]);
                    if (relationship["Type"].ends_with("/styles")) styles_path = path;
                    else if (relationship["Type"].ends_with("/theme")) theme_path = path;
                    targets[std::string{relationship["Id"]}] = std::move(path);
                }
            }
            const auto styles = styles_of(book, styles_path, theme_colors(book, theme_path));
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
                    out.push_back(sheet_of(std::string{each["name"]}, book, target == targets.end() ? std::string{} : target->second, shared, styles));
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

        // The styles a workbook's cells use, numbered as they're met (0 is the plain one), and the fonts
        // and fills they're made of: styles.xml.
        class style_table {
        public:
            // The style's number, added when it's new.
            std::size_t index(const style &look) {
                if (look.empty()) return 0;
                return position(styles_, look) + 1;
            }

            [[nodiscard]] std::string xml() {
                std::vector<font> fonts{font{}};
                std::vector<excel::borders> borders{excel::borders{}};
                std::vector<std::string> formats; // custom ones, from 164
                std::vector<sc::color> fills;
                std::string xfs = R"(<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>)";
                for (const auto &each: styles_) {
                    const auto font_id = position(fonts, each.font);
                    const auto fill_id = each.background ? position(fills, *each.background) + 2 : 0; // after none and gray125
                    const auto border_id = position(borders, each.borders);
                    const auto format_id = format_number(formats, each.number_format);
                    std::string aligned;
                    if (each.alignment != alignment{}) {
                        const auto &align = each.alignment;
                        if (align.indent > 250) throw std::invalid_argument("excel: an indent is from 0 to 250");
                        if (align.rotation < -90 || align.rotation > 90) throw std::invalid_argument("excel: a rotation is from -90 to 90 degrees");
                        aligned = "<alignment";
                        if (align.horizontal != horizontal::general) aligned += R"( horizontal=")" + std::string{name(align.horizontal)} + '"';
                        if (align.vertical != vertical::bottom) aligned += R"( vertical=")" + std::string{name(align.vertical)} + '"';
                        if (align.wrap) aligned += R"( wrapText="1")";
                        if (align.indent) aligned += R"( indent=")" + std::to_string(align.indent) + '"';
                        if (align.stacked) aligned += R"( textRotation="255")";
                        else if (align.rotation) aligned += R"( textRotation=")" + std::to_string(align.rotation > 0 ? align.rotation : 90 - align.rotation) + '"';
                        aligned += "/>";
                    }
                    xfs += R"(<xf numFmtId=")" + std::to_string(format_id) + R"(" fontId=")" + std::to_string(font_id) + R"(" fillId=")" +
                           std::to_string(fill_id) + R"(" borderId=")" + std::to_string(border_id) + R"(" xfId="0")" +
                           (format_id ? R"( applyNumberFormat="1")" : "") + (font_id ? R"( applyFont="1")" : "") +
                           (fill_id ? R"( applyFill="1")" : "") + (border_id ? R"( applyBorder="1")" : "") +
                           (aligned.empty() ? "/>" : R"( applyAlignment="1">)" + aligned + "</xf>");
                }
                std::string xml{xml_declaration};
                xml += R"(<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">)";
                if (!formats.empty()) {
                    xml += R"(<numFmts count=")" + std::to_string(formats.size()) + R"(">)";
                    for (std::size_t i = 0; i < formats.size(); ++i)
                        xml += R"(<numFmt numFmtId=")" + std::to_string(164 + i) + R"(" formatCode=")" + xml_text(formats[i]) + R"("/>)";
                    xml += "</numFmts>";
                }
                xml += R"(<fonts count=")" + std::to_string(fonts.size()) + R"(">)";
                for (const auto &each: fonts) {
                    if (each.size && !(each.size->points() >= 1 && each.size->points() <= 409))
                        throw std::invalid_argument("excel: a font's size is from 1 to 409 points");
                    xml += "<font>";
                    if (each.bold) xml += "<b/>";
                    if (each.italic) xml += "<i/>";
                    if (each.strike) xml += "<strike/>";
                    if (each.underline == underline::single) xml += "<u/>";
                    else if (each.underline != underline::none) xml += R"(<u val=")" + std::string{name(each.underline)} + R"("/>)";
                    xml += R"(<sz val=")" + number_text(each.size ? each.size->points() : 11) + R"("/>)";
                    if (each.color) xml += R"(<color rgb=")" + argb(*each.color) + R"("/>)";
                    xml += R"(<name val="Calibri"/><family val="2"/><scheme val="minor"/></font>)";
                }
                xml += R"(</fonts><fills count=")" + std::to_string(fills.size() + 2) +
                       R"("><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>)";
                for (const auto &each: fills)
                    xml += R"(<fill><patternFill patternType="solid"><fgColor rgb=")" + argb(each) + R"("/><bgColor indexed="64"/></patternFill></fill>)";
                xml += R"(</fills><borders count=")" + std::to_string(borders.size()) + R"(">)";
                for (const auto &each: borders) {
                    const auto side = [](const char *tag, const edge &edge) {
                        if (edge.line == line::none) return "<" + std::string{tag} + "/>";
                        return "<" + std::string{tag} + R"( style=")" + std::string{name(edge.line)} + R"(">)" +
                               (edge.color ? R"(<color rgb=")" + argb(*edge.color) + R"("/>)" : std::string{}) + "</" + tag + ">";
                    };
                    const auto &up = each.diagonal_up, &down = each.diagonal_down;
                    const bool upward = up.line != line::none, downward = down.line != line::none;
                    if (upward && downward && up != down) throw std::invalid_argument("excel: both diagonals are drawn alike, so they must be the same edge");
                    xml += std::string{"<border"} + (upward ? R"( diagonalUp="1")" : "") + (downward ? R"( diagonalDown="1")" : "") + ">" +
                           side("left", each.left) + side("right", each.right) + side("top", each.top) + side("bottom", each.bottom) +
                           side("diagonal", upward ? up : down) + "</border>";
                }
                xml += R"(</borders>)"
                       R"(<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count=")" +
                       std::to_string(styles_.size() + 1) + R"(">)" + xfs +
                       R"(</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>)";
                return xml;
            }

        private:
            // Where what is in list, added at its end when it isn't there.
            template<typename T>
            static std::size_t position(std::vector<T> &list, const T &what) {
                const auto found = std::find(list.begin(), list.end(), what);
                if (found != list.end()) return static_cast<std::size_t>(found - list.begin());
                list.push_back(what);
                return list.size() - 1;
            }

            // A format code's numFmtId: 0 for General, Excel's own number for one it has without saying,
            // else the next from 164 (formats holds those).
            static std::size_t format_number(std::vector<std::string> &formats, const std::string &code) {
                if (code.empty() || code == "General") return 0;
                if (code.size() > 255) throw std::invalid_argument("excel: a number format is at most 255 characters");
                for (const auto &[id, builtin]: builtin_formats)
                    if (id != 14 && builtin == code) return id;
                return 164 + position(formats, code);
            }

            // "FFC00000": opaque, as Excel writes colours.
            static std::string argb(const sc::color &color) {
                auto hex = color.to_hex().substr(1);
                for (auto &c: hex) c = static_cast<char>(std::toupper(static_cast<unsigned char>(c)));
                return "FF" + hex;
            }

            std::vector<style> styles_;
        };

        // A cell's XML: its style's number, its formula (without the =) and the value it holds.
        std::string cell_xml(const cell &each, style_table &styles) {
            std::string formula;
            if (!each.formula.empty())
                formula = "<f>" + xml_text(each.formula.starts_with('=') ? std::string_view{each.formula}.substr(1) : std::string_view{each.formula}) + "</f>";
            auto start = R"(<c r=")" + each.name() + '"';
            if (const auto index = styles.index(each.style)) start += R"( s=")" + std::to_string(index) + '"';
            if (const auto *text = std::get_if<std::string>(&each.value)) {
                if (!formula.empty()) return start + R"( t="str">)" + formula + "<v>" + xml_text(*text) + "</v></c>";
                return start + R"( t="inlineStr"><is><t xml:space="preserve">)" + xml_text(*text) + "</t></is></c>";
            }
            if (const auto *truth = std::get_if<bool>(&each.value)) return start + R"( t="b">)" + formula + "<v>" + (*truth ? "1" : "0") + "</v></c>";
            if (const auto *number = std::get_if<double>(&each.value)) return start + ">" + formula + "<v>" + number_text(*number) + "</v></c>";
            return start + ">" + formula + "</c>";
        }

        std::string sheet_xml(const sheet &each, style_table &styles) {
            std::string xml{xml_declaration};
            xml += R"(<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">)";
            if (const auto &frozen = each.frozen()) { // the rows above and columns left of it kept in view
                const auto rows = frozen->row, columns = frozen->column;
                const auto pane = rows && columns ? "bottomRight" : rows ? "bottomLeft" : "topRight";
                xml += R"(<sheetViews><sheetView workbookViewId="0"><pane)";
                if (columns) xml += R"( xSplit=")" + std::to_string(columns) + '"';
                if (rows) xml += R"( ySplit=")" + std::to_string(rows) + '"';
                xml += R"( topLeftCell=")" + frozen->name() + R"(" activePane=")" + pane + R"(" state="frozen"/><selection pane=")" + pane +
                       R"("/></sheetView></sheetViews>)";
            }
            // The columns given a width or hidden; neighbours alike as one <col>.
            std::map<std::size_t, std::pair<std::optional<double>, bool>> columns;
            for (const auto &[column, width]: each.widths()) columns[column].first = width;
            for (const auto column: each.hidden_columns()) columns[column].second = true;
            if (!columns.empty()) {
                xml += "<cols>";
                for (auto at = columns.begin(); at != columns.end();) {
                    auto last = at;
                    while (std::next(last) != columns.end() && std::next(last)->first == last->first + 1 && std::next(last)->second == at->second) ++last;
                    const auto &[width, hidden] = at->second;
                    // A hidden column without a width of its own gets Excel's default, for when it's shown.
                    xml += R"(<col min=")" + std::to_string(at->first + 1) + R"(" max=")" + std::to_string(last->first + 1) + R"(" width=")" +
                           (width ? number_text(*width) + R"(" customWidth="1")" : std::string{R"(9.140625")"}) + (hidden ? R"( hidden="1")" : "") + "/>";
                    at = std::next(last);
                }
                xml += "</cols>";
            }
            xml += "<sheetData>";
            // The rows with cells, a height or hidden, in order.
            std::map<std::size_t, std::string> rows;
            for (const auto &row: each) {
                std::string cells;
                for (const auto &cell: row)
                    if (!cell.empty() || !cell.style.empty()) cells += cell_xml(cell, styles);
                if (!cells.empty()) rows[row.index()] = std::move(cells);
            }
            for (const auto &[index, height]: each.heights()) rows.try_emplace(index);
            for (const auto index: each.hidden_rows()) rows.try_emplace(index);
            for (const auto &[index, cells]: rows) {
                xml += R"(<row r=")" + std::to_string(index + 1) + '"';
                if (const auto height = each.height(index)) xml += R"( ht=")" + number_text(height->points()) + R"(" customHeight="1")";
                if (each.hidden_row(index)) xml += R"( hidden="1")";
                xml += cells.empty() ? "/>" : ">" + cells + "</row>";
            }
            xml += "</sheetData>";
            if (!each.merged().empty()) {
                xml += R"(<mergeCells count=")" + std::to_string(each.merged().size()) + R"(">)";
                for (const auto &cells: each.merged()) xml += R"(<mergeCell ref=")" + cells.name() + R"("/>)";
                xml += "</mergeCells>";
            }
            return xml + "</worksheet>";
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
        // The sheets first: the styles their cells use make styles.xml.
        style_table styles;
        std::vector<std::string> sheet_parts;
        for (const auto &each: book) sheet_parts.push_back(sheet_xml(each, styles));
        for (std::size_t i = 1; i <= book.size(); ++i) {
            const auto n = std::to_string(i);
            types += R"(<Override PartName="/xl/worksheets/sheet)" + n +
                     R"(.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>)";
            sheets += R"(<sheet name=")" + xml_text(names[i - 1]) + R"(" sheetId=")" + n + R"(" r:id="rId)" + n + R"("/>)";
            rels += R"(<Relationship Id="rId)" + n + R"(" Type=")" + std::string{relationships} + R"(/worksheet" Target="worksheets/sheet)" + n + R"(.xml"/>)";
        }
        const auto styles_id = "rId" + std::to_string(book.size() + 1);
        types += R"(<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>)";
        rels += R"(<Relationship Id=")" + styles_id + R"(" Type=")" + std::string{relationships} + R"(/styles" Target="styles.xml"/>)";
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
        parts.emplace_back("xl/styles.xml", styles.xml());
        for (std::size_t i = 0; i < sheet_parts.size(); ++i) parts.emplace_back("xl/worksheets/sheet" + std::to_string(i + 1) + ".xml", std::move(sheet_parts[i]));

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
