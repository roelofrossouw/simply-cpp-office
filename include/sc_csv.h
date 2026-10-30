#pragma once

#include "csv_cell.h"
#include "csv_row.h"
#include "csv_sheet.h"
#include "excel_range.h"

#include <filesystem>
#include <string>
#include <string_view>

// CSV files (RFC 4180): a sheet of values, read and written with the same access as an Excel
// sheet (excel.h), so code that only uses data works on both. No styles, widths or formulas.
//
//     const auto prices = sc::office::csv::read("prices.csv");          // , or ; found by itself
//     for (const auto &row: prices)
//         for (const auto &cell: row)
//             cell.row, cell.column, cell.value;                         // 3, 4, 12.5
//     if (prices.contains("B2")) total = std::get<double>(prices.at("B2").value);
//
//     sc::office::csv::sheet out;
//     out["A1"] = "Item";
//     out[{1, 1}] = 12.5;                                                // B2
//     sc::office::csv::write(out, "out.csv");
//
// Fields with the delimiter, a quote or a line break are quoted, a quote in them doubled ("He said
// ""hi"""), and a quoted field may hold line breaks.
namespace sc::office::csv {
    using range = excel::range;

    struct options {
        // Reading: the delimiter, or '\0' to find it: a comma, unless the file reads more evenly
        // with semicolons (as Excel writes CSV where the decimal separator is a comma).
        // Writing: '\0' writes commas.
        char delimiter = '\0';
        // Reading: plain numbers ("12.5", "-3", "1e6", not "007" or "+1") as numbers and TRUE or
        // FALSE (any case) as true/false, as Excel reads them; a quoted field stays text. false
        // reads every field as text.
        bool typed = true;
        // Writing: a UTF-8 byte order mark first, which Excel needs to read UTF-8 text as such.
        // Reading skips one whichever way.
        bool byte_order_mark = false;
    };

    // Reads a CSV file into a sheet named after the file. Empty fields are cells that aren't set,
    // so rows keep their places. Throws std::runtime_error when it can't be read, or for a quoted
    // field that isn't closed.
    [[nodiscard]] sheet read(const std::filesystem::path &file, const options &how = {});
    // The same, for CSV text held in memory; the sheet has no name.
    [[nodiscard]] sheet read_from(std::string_view text, const options &how = {});

    // The delimiter read() would find for text: ',' or ';'.
    [[nodiscard]] char find_delimiter(std::string_view text);

    // Writes the sheet's values as CSV: from A1 to its last row and column, CRLF line ends.
    // Numbers are written as briefly as they read back the same, true/false as TRUE and FALSE, and
    // text that would read back as a number or TRUE/FALSE is quoted, so it stays text. Throws
    // std::invalid_argument for a number that isn't finite, std::runtime_error when the file
    // can't be written.
    void write(const sheet &values, const std::filesystem::path &file, const options &how = {});
    // An Excel sheet's values, the same way.
    void write(const excel::sheet &values, const std::filesystem::path &file, const options &how = {});
    // What write() writes, as text.
    [[nodiscard]] std::string bytes(const sheet &values, const options &how = {});
    [[nodiscard]] std::string bytes(const excel::sheet &values, const options &how = {});
}
