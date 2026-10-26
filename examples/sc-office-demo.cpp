// Lists the pictures in a Word document, in the order they appear, or saves them with --save; prints
// an Excel workbook's cells, or writes a small one with --write.
//
//     sc-office-demo statement.docx
//     sc-office-demo statement.docx --save pictures
//     sc-office-demo prices.xlsx
//     sc-office-demo --write sample.xlsx

#include <office.h>
#include <params.h>

#include <filesystem>
#include <fstream>
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

    void pictures(const sc::params &args) {
        // [readme]
        const auto pictures = sc::office::word::images(args.as<std::string>("file"));
        const std::filesystem::path folder = args.as<std::string>("save");
        if (!folder.empty()) std::filesystem::create_directories(folder);
        for (const auto &picture: pictures) {
            std::cout << picture.name << "  " << picture.bytes.size() << " bytes" << (picture.raster() ? "" : "  (a drawing)") << '\n';
            if (!folder.empty()) std::ofstream{folder / picture.name, std::ios::binary} << picture.bytes;
        }
        std::cout << pictures.size() << " pictures\n";
        // [/readme]
    }

    void cells(const std::string &file) {
        for (const auto &sheet: sc::office::excel::read(file)) {
            std::cout << "== " << sheet.name() << '\n';
            for (const auto &row: sheet)
                for (const auto &cell: row)
                    std::cout << cell.name() << '\t' << text_of(cell.value) << (cell.formula.empty() ? "" : "\t" + cell.formula) << '\n';
        }
    }

    void sample(const std::string &file) {
        sc::office::excel::workbook book;
        auto &prices = book["Prices"];
        prices["A1"] = "Item";
        prices["B1"] = "Price";
        prices["A2"] = "Bread";
        prices["B2"] = 18.5;
        prices["A3"] = "Milk";
        prices["B3"] = 21.0;
        prices["A4"] = "Total";
        prices["B4"].formula = "=SUM(B2:B3)";
        sc::office::excel::write(book, file);
        std::cout << "wrote " << file << '\n';
    }
}

int main(int argc, char **argv) {
    const sc::params args{argc, argv, {
        {"name", "sc-office-demo"},
        {"description", "Lists the pictures in a Word document (.docx): those in its text in the order they appear, then "
                        "the rest; with --save, writes them into a folder as they're stored. Prints the cells of an Excel "
                        "workbook (.xlsx), with their values and formulas; with --write, writes a small one instead."},
        {"options", {
            {"save", {{"short", "s"}, {"value", "FOLDER"}, {"default", ""}, {"help", "write a .docx's pictures into FOLDER"}}},
            {"write", {{"short", "w"}, {"type", "boolean"}, {"help", "write a small workbook to FILE (.xlsx)"}}}}},
        {"positional", {
            {"file", {{"help", "the .docx or .xlsx"}}}}}}};
    try {
        const auto file = args.as<std::string>("file");
        if (args["write"].get<bool>()) sample(file);
        else if (std::filesystem::path{file}.extension() == ".xlsx") cells(file);
        else pictures(args);
    } catch (const std::exception &error) {
        std::cerr << "sc-office-demo: " << error.what() << '\n';
        return 1;
    }
    return 0;
}
