// CSV: quoting (RFC 4180), line breaks in fields, delimiters found or given, typed values, empty
// fields, line ends and byte order marks, written and read back, and the same code working on a
// CSV sheet and an Excel sheet.
#include <sc_csv.h>
#include <excel.h>
#include <sc_test.h>

#include <filesystem>
#include <fstream>
#include <limits>
#include <random>
#include <string>

namespace fs = std::filesystem;
using sc::office::csv::value;

namespace {
    // Code that only uses data: the same for either kind of sheet.
    template<typename any_sheet>
    double total(const any_sheet &sheet, const std::size_t column) {
        double sum = 0;
        for (const auto &row: sheet)
            for (const auto &cell: row)
                if (cell.column == column && std::holds_alternative<double>(cell.value)) sum += std::get<double>(cell.value);
        return sum;
    }

    template<typename any_sheet>
    void fill(any_sheet &sheet) {
        sheet["A1"] = "Item";
        sheet["B1"] = "Price";
        sheet[{1, 0}] = "Bread";
        sheet[{1, 1}] = 12.5;
        sheet["A3"] = "Milk, full cream";
        sheet["B3"] = 20.0;
        sheet["C3"] = true;
    }
}

int main() {
    using namespace sc::office::csv;
    const auto root = fs::temp_directory_path() / ("sc-office-csv-" + std::to_string(std::random_device{}()));
    fs::create_directories(root);

    SECTION("Quoting: delimiters, quotes and line breaks in a field, as RFC 4180 says");
    {
        const auto sheet = read_from("name,said\r\n\"Smith, J\",\"He said \"\"hi\"\"\"\r\n\"two\nlines\",x\r\n");
        CHECK_EQ(std::get<std::string>(sheet.at("A2").value), std::string{"Smith, J"});
        CHECK_EQ(std::get<std::string>(sheet.at("B2").value), std::string{"He said \"hi\""});
        CHECK_EQ(std::get<std::string>(sheet.at("A3").value), std::string{"two\nlines"});
        CHECK_EQ(sheet.size(), std::size_t{3});
        CHECK_THROWS_AS((void) read_from("a,\"not closed\nb,c\n"), std::runtime_error);
    }

    SECTION("Line ends (CRLF, LF, CR), a byte order mark, blank lines and empty fields");
    {
        for (const std::string end: {"\r\n", "\n", "\r"}) {
            const auto sheet = read_from("\xEF\xBB\xBF" "a,b" + end + end + "c,,d" + end);
            CHECK_EQ(std::get<std::string>(sheet.at("A1").value), std::string{"a"}); // the mark isn't part of it
            CHECK(!sheet.contains("A2"));                                             // a blank line keeps its place
            CHECK_EQ(std::get<std::string>(sheet.at("C3").value), std::string{"d"});
            CHECK(!sheet.contains("B3"));                                             // an empty field is no cell
        }
        CHECK(read_from("").empty());
    }

    SECTION("The delimiter: found (comma, or semicolons when they fit better), or given");
    {
        CHECK_EQ(find_delimiter("a,b,c\n1,2,3\n"), ',');
        CHECK_EQ(find_delimiter("Name;Price\nBread;12,50\nMilk;20,00\n"), ';');  // decimal commas
        CHECK_EQ(find_delimiter("Name;Price;Note\nBread;12;a,b\n"), ';');
        CHECK_EQ(find_delimiter("just\none\ncolumn\n"), ',');                    // nothing better: a comma
        CHECK_EQ(find_delimiter("a;b\nc\n"), ',');                               // ; uneven too: a comma
        const auto european = read_from("Name;Price\nBread;12,50\n");
        CHECK_EQ(std::get<std::string>(european.at("B2").value), std::string{"12,50"});
        const auto tabs = read_from("a\tb\n1\t2\n", {.delimiter = '\t'});
        CHECK_EQ(std::get<double>(tabs.at("B2").value), 2.0);
    }

    SECTION("Typed values: plain numbers and TRUE/FALSE, quoted ones stay text; or everything text");
    {
        const auto sheet = read_from("12.5,-3,1e6,007,+1,.5,TRUE,false,\"12\",0\n");
        CHECK_EQ(std::get<double>(sheet.at("A1").value), 12.5);
        CHECK_EQ(std::get<double>(sheet.at("B1").value), -3.0);
        CHECK_EQ(std::get<double>(sheet.at("C1").value), 1e6);
        CHECK_EQ(std::get<std::string>(sheet.at("D1").value), std::string{"007"});
        CHECK_EQ(std::get<std::string>(sheet.at("E1").value), std::string{"+1"});
        CHECK_EQ(std::get<std::string>(sheet.at("F1").value), std::string{".5"});
        CHECK(std::get<bool>(sheet.at("G1").value));
        CHECK(!std::get<bool>(sheet.at("H1").value));
        CHECK_EQ(std::get<std::string>(sheet.at("I1").value), std::string{"12"});
        CHECK_EQ(std::get<double>(sheet.at("J1").value), 0.0);
        const auto text = read_from("12.5,TRUE\n", {.typed = false});
        CHECK_EQ(std::get<std::string>(text.at("A1").value), std::string{"12.5"});
        CHECK_EQ(std::get<std::string>(text.at("B1").value), std::string{"TRUE"});
    }

    SECTION("Writing: quoted when needed, text kept text, read back the same");
    {
        sheet out;
        fill(out);
        out["D4"] = "12";                    // text that looks like a number
        out["A5"] = "He said \"hi\"\nthen left";
        const auto text = bytes(out);
        CHECK(text.starts_with("Item,Price,,\r\nBread,12.5,,\r\n\"Milk, full cream\",20,TRUE,\r\n,,,\"12\"\r\n"));
        CHECK(text.ends_with("\"He said \"\"hi\"\"\nthen left\",,,\r\n"));
        const auto back = read_from(text);
        for (const auto &row: out)
            for (const auto &cell: row) CHECK(back.at({cell.row, cell.column}).value == cell.value);
        CHECK_EQ(bytes(out, {.delimiter = ';'}).substr(0, 15), std::string{"Item;Price;;\r\nB"});
        CHECK(bytes(out, {.byte_order_mark = true}).starts_with("\xEF\xBB\xBF" "Item"));
        // Numbers as other programs read them: no exponent for whole or everyday numbers
        sheet numbers;
        numbers["A1"] = 1e6;
        numbers["B1"] = 123456789012.0;
        numbers["C1"] = -2500000.0;
        numbers["D1"] = 0.000125;
        numbers["E1"] = 1234567.5;
        numbers["F1"] = 0.1;
        numbers["G1"] = 1e25;
        numbers["H1"] = 1e-9;
        numbers["I1"] = 0.0;
        CHECK_EQ(bytes(numbers), std::string{"1000000,123456789012,-2500000,0.000125,1234567.5,0.1,1e+25,1e-09,0\r\n"});
        const auto read_numbers = read_from(bytes(numbers));
        for (const auto &row: numbers)
            for (const auto &cell: row) CHECK(read_numbers.at({cell.row, cell.column}).value == cell.value);
        sheet bad;
        bad["A1"] = std::numeric_limits<double>::infinity();
        CHECK_THROWS_AS((void) bytes(bad), std::invalid_argument);
        CHECK(bytes(sheet{}).empty());
    }

    SECTION("Access as an Excel sheet's: names, places, at, contains, erase, iteration");
    {
        sheet values{"prices"};
        fill(values);
        CHECK_EQ(values.name(), std::string{"prices"});
        CHECK(values.contains("b2") && !values.contains("Z9"));
        CHECK_THROWS_AS((void) values.at("Z9"), std::out_of_range);
        CHECK_EQ(values.at("A3").name(), std::string{"A3"});
        CHECK_EQ(values.erase("C3"), std::size_t{1});
        CHECK_EQ(values.erase("C3"), std::size_t{0});
        CHECK_EQ(values.size(), std::size_t{3});
        CHECK_EQ(values.at("A2").row, std::size_t{1});
    }

    SECTION("The same code on a CSV sheet and an Excel sheet, and one turned into the other");
    {
        sheet csv_values;
        fill(csv_values);
        sc::office::excel::workbook book;
        auto &excel_values = book["Prices"];
        fill(excel_values);
        CHECK_EQ(total(csv_values, 1), 32.5);
        CHECK_EQ(total(excel_values, 1), 32.5);

        excel_values["B4"].formula = "=SUM(B2:B3)";
        excel_values["B4"].value = 32.5;     // what Excel calculated
        excel_values["A1"].style.font.bold = true;
        const sheet from_excel{excel_values};
        CHECK_EQ(from_excel.name(), std::string{"Prices"});
        CHECK_EQ(std::get<double>(from_excel.at("B4").value), 32.5); // the formula's value
        CHECK_EQ(bytes(excel_values), bytes(from_excel));            // an Excel sheet writes its values
        const auto to_excel = csv_values.to_excel();
        CHECK_EQ(to_excel.name(), std::string{"Sheet1"});
        CHECK(std::get<bool>(to_excel.at("C3").value));
    }

    SECTION("Files: read() names the sheet after the file");
    {
        sheet out;
        fill(out);
        write(out, root / "prices.csv");
        const auto back = read(root / "prices.csv");
        CHECK_EQ(back.name(), std::string{"prices"});
        CHECK_EQ(total(back, 1), 32.5);
        CHECK_THROWS_AS((void) read(root / "missing.csv"), std::runtime_error);
        CHECK_THROWS_AS(write(out, root / "no" / "such" / "folder.csv"), std::runtime_error);
    }

    fs::remove_all(root);
    TEST_SUMMARY();
}
