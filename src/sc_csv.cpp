#include "sc_csv.h"

#include <algorithm>
#include <cctype>
#include <charconv>
#include <cmath>
#include <fstream>
#include <iterator>
#include <stdexcept>
#include <vector>

namespace sc::office::csv {
    namespace {
        constexpr std::string_view byte_order_mark = "\xEF\xBB\xBF";

        struct field {
            std::string text;
            bool quoted = false;
        };
        using record = std::vector<field>;

        // RFC 4180 records: fields split by delimiter, a quoted field's "" a quote and its line
        // breaks kept; lines end with \n, \r\n or \r. limit: stop after so many records (0: all).
        std::vector<record> parse(std::string_view text, const char delimiter, const std::size_t limit = 0) {
            if (text.starts_with(byte_order_mark)) text.remove_prefix(byte_order_mark.size());
            std::vector<record> records;
            std::size_t i = 0;
            while (i < text.size() && (limit == 0 || records.size() < limit)) {
                record fields;
                while (true) {
                    field item;
                    if (i < text.size() && text[i] == '"') {
                        item.quoted = true;
                        ++i;
                        while (true) {
                            if (i >= text.size()) throw std::runtime_error("csv: a quoted field isn't closed (row " + std::to_string(records.size() + 1) + ")");
                            if (text[i] == '"') {
                                if (i + 1 < text.size() && text[i + 1] == '"') {
                                    item.text += '"';
                                    i += 2;
                                    continue;
                                }
                                ++i;
                                break;
                            }
                            item.text += text[i++];
                        }
                    }
                    // Unquoted text, or anything after a closing quote, up to the delimiter or line end.
                    while (i < text.size() && text[i] != delimiter && text[i] != '\n' && text[i] != '\r') item.text += text[i++];
                    fields.push_back(std::move(item));
                    if (i < text.size() && text[i] == delimiter) {
                        ++i;
                        continue;
                    }
                    break;
                }
                if (i < text.size() && text[i] == '\r') ++i;
                if (i < text.size() && text[i] == '\n') ++i;
                records.push_back(std::move(fields));
            }
            return records;
        }

        // Whether records split on a delimiter look like a table: more than one field, and the
        // same number in each (blank lines aside).
        bool even(const std::vector<record> &records) {
            std::size_t fields = 0;
            for (const auto &each: records) {
                if (each.size() == 1 && each[0].text.empty() && !each[0].quoted) continue;
                if (fields == 0) fields = each.size();
                else if (each.size() != fields) return false;
            }
            return fields > 1;
        }

        // "12.5", "-3", "1e6", "0.5": a plain number, as Excel reads one; not "007", "+1" or ".5".
        bool plain_number(const std::string_view text) {
            std::size_t i = 0;
            const auto digits = [&] {
                const auto start = i;
                while (i < text.size() && std::isdigit(static_cast<unsigned char>(text[i]))) ++i;
                return i - start;
            };
            if (i < text.size() && text[i] == '-') ++i;
            const auto start = i;
            const auto whole = digits();
            if (whole == 0 || (whole > 1 && text[start] == '0')) return false;
            if (i < text.size() && text[i] == '.') {
                ++i;
                if (digits() == 0) return false;
            }
            if (i < text.size() && (text[i] == 'e' || text[i] == 'E')) {
                ++i;
                if (i < text.size() && (text[i] == '+' || text[i] == '-')) ++i;
                if (digits() == 0) return false;
            }
            return i == text.size();
        }

        std::string lowercase(std::string text) {
            std::ranges::transform(text, text.begin(), [](const unsigned char c) { return static_cast<char>(std::tolower(c)); });
            return text;
        }

        value typed_value(field item, const bool typed) {
            if (!typed || item.quoted) return std::move(item.text);
            if (plain_number(item.text)) {
                double number = 0;
                const auto [end, error] = std::from_chars(item.text.data(), item.text.data() + item.text.size(), number);
                if (error == std::errc{} && std::isfinite(number)) return number;
            }
            if (const auto word = lowercase(item.text); word == "true" || word == "false") return word == "true";
            return std::move(item.text);
        }

        std::string number_text(const double number) {
            if (!std::isfinite(number)) throw std::invalid_argument("csv: a number that isn't finite");
            char text[32];
            const auto [end, error] = std::to_chars(text, text + sizeof text, number);
            return {text, end};
        }

        // A value as a field: quoted when it holds the delimiter, a quote or a line break, or is
        // text that would read back as a number or TRUE/FALSE.
        std::string field_text(const value &item, const char delimiter) {
            if (const auto *number = std::get_if<double>(&item)) return number_text(*number);
            if (const auto *flag = std::get_if<bool>(&item)) return *flag ? "TRUE" : "FALSE";
            const auto *text = std::get_if<std::string>(&item);
            if (!text || text->empty()) return {};
            const auto word = lowercase(*text);
            if (text->find_first_of(std::string{delimiter} + "\"\r\n") == std::string::npos && !plain_number(*text) && word != "true" &&
                word != "false")
                return *text;
            std::string quoted = "\"";
            for (const char c: *text) quoted += c == '"' ? std::string{"\"\""} : std::string(1, c);
            return quoted + '"';
        }

        // Any sheet with rows of cells holding values: csv's or Excel's.
        template<typename any_sheet>
        std::string written(const any_sheet &values, const options &how) {
            const char delimiter = how.delimiter ? how.delimiter : ',';
            std::size_t last_row = 0, columns = 0;
            for (const auto &each: values)
                for (const auto &item: each) {
                    if (std::holds_alternative<std::monostate>(item.value)) continue;
                    last_row = std::max(last_row, item.row + 1);
                    columns = std::max(columns, item.column + 1);
                }
            std::string out{how.byte_order_mark ? byte_order_mark : std::string_view{}};
            auto each = values.begin();
            for (std::size_t row = 0; row < last_row; ++row) {
                while (each != values.end() && each->index() < row) ++each;
                std::vector<std::string> fields(columns);
                if (each != values.end() && each->index() == row)
                    for (const auto &item: *each) fields[item.column] = field_text(item.value, delimiter);
                for (std::size_t column = 0; column < columns; ++column) out += (column ? std::string(1, delimiter) : std::string{}) + fields[column];
                out += "\r\n";
            }
            return out;
        }

        void save(const std::string &text, const std::filesystem::path &file) {
            std::ofstream out{file, std::ios::binary | std::ios::trunc};
            out << text;
            if (!out) throw std::runtime_error("csv: can't write " + file.string());
        }
    }

    char find_delimiter(const std::string_view text) {
        const auto sample = [&](const char delimiter) {
            try {
                return even(parse(text, delimiter, 50));
            } catch (const std::runtime_error &) {
                return false;
            }
        };
        if (sample(',')) return ',';
        return sample(';') ? ';' : ',';
    }

    sheet read_from(const std::string_view text, const options &how) {
        const char delimiter = how.delimiter ? how.delimiter : find_delimiter(text);
        sheet out;
        auto records = parse(text, delimiter);
        for (std::size_t row = 0; row < records.size(); ++row)
            for (std::size_t column = 0; column < records[row].size(); ++column) {
                auto &item = records[row][column];
                if (item.text.empty()) continue;
                out[{row, column}] = typed_value(std::move(item), how.typed);
            }
        return out;
    }

    sheet read(const std::filesystem::path &file, const options &how) {
        std::ifstream in{file, std::ios::binary};
        if (!in) throw std::runtime_error("csv: can't read " + file.string());
        const std::string text{std::istreambuf_iterator<char>{in}, std::istreambuf_iterator<char>{}};
        auto values = read_from(text, how);
        sheet named{file.stem().string()};
        for (const auto &each: values)
            for (const auto &item: each) named[{item.row, item.column}] = item.value;
        return named;
    }

    std::string bytes(const sheet &values, const options &how) { return written(values, how); }
    std::string bytes(const excel::sheet &values, const options &how) { return written(values, how); }
    void write(const sheet &values, const std::filesystem::path &file, const options &how) { save(written(values, how), file); }
    void write(const excel::sheet &values, const std::filesystem::path &file, const options &how) { save(written(values, how), file); }
}
