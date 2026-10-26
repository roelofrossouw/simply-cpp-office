// Prints the cells of an Excel workbook, sheet by sheet, with their values and formulas, or writes
// a small one with --write: a header in bold white on blue, widened columns, a total in bold between
// a thin line and a double one, as accountants rule them, prices to two decimals, a taller header
// row, and an italic, underlined note merged across both columns.
//
//     sc-office-excel-demo prices.xlsx
//     sc-office-excel-demo --write sample.xlsx

#include <excel.h>
#include <params.h>

#include <iostream>
#include <sstream>

namespace {
    // A value as text: numbers as C++ prints them, TRUE/FALSE, text as it is.
    std::string text_of(const sc::office::excel::value &value) {
        if (const auto *text = std::get_if<std::string>(&value)) return *text;
        if (const auto *number = std::get_if<double>(&value)) {
            std::ostringstream out;
            out << *number;
            return out.str();
        }
        if (const auto *truth = std::get_if<bool>(&value)) return *truth ? "TRUE" : "FALSE";
        return {};
    }

    void sample(const std::string &file) {
        using namespace sc::office::excel;
        workbook book;
        auto &prices = book["Prices"];
        const style header{.font = {.size = 12, .color = sc::color::White, .bold = true}, .background = sc::color{"#4472c4"},
                           .alignment = {.horizontal = horizontal::center}};
        prices["A1"] = "Item";
        prices["B1"] = "Price";
        prices["A1"].style = prices["B1"].style = header;
        prices["A2"] = "Bread";
        prices["B2"] = 18.5;
        prices["A3"] = "Milk";
        prices["B3"] = 21.0;
        prices["A4"] = "Total";
        prices["B4"].formula = "=SUM(B2:B3)";
        prices["A4"].style.font.bold = prices["B4"].style.font.bold = true;
        prices["B4"].style.font.color = sc::color{"#c00000"};
        for (const auto *cell: {"A4", "B4"}) {
            prices[cell].style.borders.top = {line::thin};
            prices[cell].style.borders.bottom = {line::doubled};
        }
        prices["A6"] = "Prices include VAT";
        prices["A6"].style.font.italic = true;
        prices["A6"].style.font.underline = underline::single;
        prices.merge("A6:B6");
        prices.height(0, 24); // points
        for (const auto *cell: {"B2", "B3", "B4"}) prices[cell].style.number_format = "#,##0.00";
        prices.width("A", 20);
        prices.width("B", 12);
        write(book, file);
        std::cout << "wrote " << file << '\n';
    }
}

int main(int argc, char **argv) {
    const sc::params args{argc, argv, {
        {"name", "sc-office-excel-demo"},
        {"description", "Prints the cells of an Excel workbook (.xlsx), sheet by sheet, with their values and formulas; with "
                        "--write, writes a small workbook to the file instead."},
        {"options", {
            {"write", {{"short", "w"}, {"type", "boolean"}, {"help", "write a small workbook to FILE"}}}}},
        {"positional", {
            {"file", {{"help", "the .xlsx"}}}}}}};
    try {
        const auto file = args.as<std::string>("file");
        if (args["write"].get<bool>()) {
            sample(file);
            return 0;
        }
        // [readme]
        for (const auto &sheet: sc::office::excel::read(file)) {
            std::cout << "== " << sheet.name() << '\n';
            for (const auto &row: sheet)
                for (const auto &cell: row)
                    std::cout << cell.name() << '\t' << text_of(cell.value) << (cell.formula.empty() ? "" : "\t" + cell.formula) << '\n';
        }
        // [/readme]
    } catch (const std::exception &error) {
        std::cerr << "sc-office-excel-demo: " << error.what() << '\n';
        return 1;
    }
    return 0;
}
