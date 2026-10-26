#pragma once

#include "excel_address.h"
#include "excel_cell.h"
#include "excel_range.h"
#include "excel_row.h"
#include "excel_sheet.h"
#include "excel_style.h"
#include "excel_workbook.h"

#include <filesystem>
#include <string>

// Excel workbooks (.xlsx): sheets of cells, each with its value and formula, read and written. A
// workbook is a ZIP archive of XML (Office Open XML), read and written with sc-archive and sc-web.
// Cells have a style (font size, colour, bold, italic and underline; a background; borders;
// alignment; a number format), columns a width (or are hidden), rows a height, cells can be merged
// and panes frozen; no pictures. Workbooks, sheets and rows work
// like the standard containers: a sheet holds the rows that have cells, a row the cells that are
// set.
//
//     const auto book = sc::office::excel::read("prices.xlsx");
//     for (const auto &sheet: book)
//         for (const auto &row: sheet)
//             for (const auto &cell: row)
//                 cell.row, cell.column, cell.value, cell.formula;   // 3, 4, "Hello", "=C6"
//     if (book.contains("Prices") && book["Prices"].contains("B2"))
//         total = std::get<double>(book["Prices"].at("B2").value);
//
//     sc::office::excel::workbook out;
//     auto &sheet = out["Prices"];            // added, as std::map would
//     sheet["A1"] = "Item";
//     sheet[{1, 1}] = 12.5;                   // B2
//     sheet["B3"].formula = "=B2*2";
//     sheet["A1"].style.font.bold = true;
//     sheet.width("A", 20);                   // characters
//     sc::office::excel::write(out, "out.xlsx");
namespace sc::office::excel {
    // Reads every sheet's cells, in the workbook's order: values, formulas with "=" (a formula
    // Excel keeps once for a range is given to each cell, its relative references moved), styles
    // (theme and indexed colours as sc::color) and the column widths set by hand. Throws
    // std::runtime_error for a file that isn't an .xlsx (not a ZIP archive, or no
    // xl/workbook.xml).
    [[nodiscard]] workbook read(const std::filesystem::path &file);
    // The same, for an .xlsx held in memory.
    [[nodiscard]] workbook read_from(std::string bytes);

    // Writes the workbook as an .xlsx: each cell's value, formula and style, text inline, and the
    // columns' widths. Excel calculates the formulas when it opens the file. Throws
    // std::invalid_argument for a workbook without sheets, a sheet name Excel doesn't take (one
    // read from a file), a number that isn't finite or a font size outside 1 to 409 points.
    void write(const workbook &book, const std::filesystem::path &file);
    // What write() writes, as bytes: to send, or keep in memory.
    [[nodiscard]] std::string bytes(const workbook &book);
}
