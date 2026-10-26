// Cell names, sheets and workbooks; a workbook written and read back; an .xlsx built here the way
// Excel writes one (shared strings, rich text, formulas, errors, a relationship's absolute target).
#include <sc_archive.h>
#include <excel.h>
#include <sc_test.h>

#include <algorithm>
#include <cmath>
#include <filesystem>
#include <fstream>
#include <iterator>
#include <limits>
#include <random>
#include <string>
#include <utility>
#include <vector>

namespace fs = std::filesystem;
using sc::office::excel::address;
using sc::office::excel::value;

namespace {
    const std::string workbook_xml = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="2" r:id="rId9"/>
    <sheet name="Data &amp; more" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>)";
    const std::string workbook_rels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>)";
    const std::string shared_strings = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Item</t></si>
  <si><r><rPr><b/></rPr><t>Bold</t></r><r><t xml:space="preserve"> and plain</t></r></si>
  <si><t>日本</t><rPh sb="0" eb="2"><t>にほん</t></rPh></si>
  <si><t>Tab_x0009_here</t></si>
</sst>)";
    const std::string sheet1 = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row>
    <row r="3"><c r="A3"><v>12.5</v></c><c r="B3" t="b"><v>1</v></c><c r="C3" t="s"><v>2</v></c></row>
    <row><c><f>A3*2</f><v>25</v></c><c t="str"><f>"x"&amp;"y"</f><v>xy</v></c><c t="e"><f>1/0</f><v>#DIV/0!</v></c></row>
    <row r="6"><c r="B6" t="inlineStr"><is><t>inline</t></is></c><c r="C6" s="1"/><c r="D6" t="s"><v>3</v></c></row>
    <row r="7"><c r="A7"><f t="shared" ref="A7:A9" si="0">A3+$B$1*C3&amp;"A1"+'My sheet'!B2+SUM(D3:D4)+LOG10(1E5)</f><v>1</v></c></row>
    <row r="8"><c r="A8"><f t="shared" si="0"/><v>2</v></c><c r="B8"><f>NOW()</f></c></row>
    <row r="9"><c r="A9"><f t="shared" si="0"/></c></row>
  </sheetData>
</worksheet>)";
    const std::string sheet2 = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"><c r="B2"><v>-3E-2</v></c></row></sheetData></worksheet>)";

    bool is_empty(const value &v) { return std::holds_alternative<std::monostate>(v); }

    std::string text(const sc::office::excel::sheet &sheet, const address where) { return std::get<std::string>(sheet.at(where).value); }
    double number(const sc::office::excel::sheet &sheet, const address where) { return std::get<double>(sheet.at(where).value); }
}

int main() {
    const auto root = fs::temp_directory_path() / ("sc-office-excel-test-" + std::to_string(std::random_device{}()));
    fs::create_directories(root);

    SECTION("Addresses, by place or name");
    {
        CHECK(address{"A1"} == (address{0, 0}));
        CHECK(address{"B3"} == (address{2, 1}));
        CHECK(address{"aa10"} == (address{9, 26}));
        CHECK(address{std::string{"XFD1048576"}} == (address{1048575, 16383}));
        CHECK_EQ((address{2, 1}.name()), std::string{"B3"});
        CHECK_EQ((address{0, 701}.name()), std::string{"ZZ1"});
        CHECK_EQ(sc::office::excel::column_name(0), std::string{"A"});
        CHECK_EQ(sc::office::excel::column_name(25), std::string{"Z"});
        CHECK_EQ(sc::office::excel::column_name(26), std::string{"AA"});
        CHECK_EQ(sc::office::excel::column_name(702), std::string{"AAA"});
        for (const auto *wrong: {"", "A", "1", "3B", "A0", "A1B", "XFE1", "A1048577", "A-1"})
            CHECK_THROWS_AS((void) address{wrong}, std::invalid_argument);
    }

    SECTION("A sheet: only the rows and cells that are set");
    {
        sc::office::excel::sheet sheet{"S"};
        CHECK(sheet.empty());
        CHECK(!sheet.contains("A1"));
        CHECK_THROWS_AS((void) sheet.at("A1"), std::out_of_range);
        sheet["E4"] = "Hello";
        sheet["C2"] = 2.0;
        sheet["A4"].formula = "=C2*2";
        CHECK(sheet.contains("C2"));
        CHECK(!sheet.contains("B2"));
        CHECK_EQ(sheet.size(), std::size_t{2}); // rows 2 and 4 only
        CHECK(text(std::as_const(sheet), "E4") == "Hello");
        std::vector<std::string> seen;
        for (const auto &row: sheet)
            for (const auto &cell: row) seen.push_back(cell.name() + "@" + std::to_string(row.index()));
        CHECK_EQ(seen.size(), std::size_t{3});
        CHECK_EQ(seen[0], std::string{"C2@1"});
        CHECK_EQ(seen[1], std::string{"A4@3"}); // a row's cells in column order
        CHECK_EQ(seen[2], std::string{"E4@3"});
        const auto &hello = sheet.at("E4");
        CHECK_EQ(hello.row, std::size_t{3});
        CHECK_EQ(hello.column, std::size_t{4});
        CHECK(hello.formula.empty());
        CHECK_EQ(sheet.at("A4").formula, std::string{"=C2*2"});
        CHECK(is_empty(sheet.at("A4").value));
        CHECK_THROWS_AS(sheet[address(1048576, 0)] = 1.0, std::out_of_range);
        // Changing cells while going through them.
        for (auto &row: sheet)
            for (auto &cell: row)
                if (const auto *n = std::get_if<double>(&cell.value)) cell = *n * 10;
        CHECK_EQ(number(sheet, "C2"), 20.0);
    }

    SECTION("A workbook: its sheets by name");
    {
        sc::office::excel::workbook book;
        CHECK(book.empty());
        auto &prices = book.add("Prices");
        CHECK_THROWS_AS(book.add("prices"), std::invalid_argument); // Excel ignores case
        CHECK_THROWS_AS(book.add(""), std::invalid_argument);
        CHECK_THROWS_AS(book.add("a/b"), std::invalid_argument);
        CHECK_THROWS_AS(book.add("'quoted'"), std::invalid_argument);
        CHECK_THROWS_AS(book.add(std::string(32, 'x')), std::invalid_argument);
        CHECK_NOTHROW(book.add(std::string(31, 'x')));
        CHECK_NOTHROW(book.add("Résumé 2026"));
        CHECK(book.contains("PRICES"));
        CHECK(!book.contains("none"));
        CHECK(&book.at("PRICES") == &prices);
        CHECK(&book["prices"] == &prices); // there already: nothing added
        CHECK_EQ(book.size(), std::size_t{3});
        auto &added = book["New"];          // added, as std::map would
        CHECK_EQ(book.size(), std::size_t{4});
        CHECK_EQ(added.name(), std::string{"New"});
        CHECK_THROWS_AS(book["a/b"], std::invalid_argument);
        CHECK_THROWS_AS((void) book.at("none"), std::out_of_range);
        CHECK_THROWS_AS((void) std::as_const(book)["none"], std::out_of_range); // a const one adds nothing
        for (int i = 0; i < 100; ++i) book.add("More " + std::to_string(i));
        prices["A1"] = "still there"; // references stay good as sheets are added
        CHECK(text(book.front(), "A1") == "still there");
        std::vector<std::string> names;
        for (const auto &sheet: book) names.push_back(sheet.name());
        CHECK_EQ(names.size(), std::size_t{104});
        CHECK_EQ(names[3], std::string{"New"});
    }

    SECTION("Written and read back");
    {
        sc::office::excel::workbook book;
        auto &prices = book["Prices"];
        prices["A1"] = "Item & <price>";
        prices["B1"] = "  spaced  ";
        prices["A2"] = "Bread";
        prices["B2"] = 12.5;
        prices["C2"] = true;
        prices["D2"] = false;
        prices["A4"] = 0.1;
        prices["B4"] = -1e300;
        prices["C4"] = "line\nbreak, bell\a, _x0041_ as it is, ünïcödé";
        prices["AB10"] = 42.0;
        prices["B5"].formula = "=B2*2";                  // never calculated
        prices["C5"] = 25.0;
        prices["C5"].formula = "=B2*2";                  // with the value it gave
        prices["D5"] = "yes";
        prices["D5"].formula = R"(=IF(C2,"yes","no"))";
        prices["E5"];                                    // added, never set: not written
        book["Empty"];
        const auto file = root / "round.xlsx";
        sc::office::excel::write(book, file);
        const auto archive = sc::archive::open(file);
        CHECK(archive.contains("xl/worksheets/sheet2.xml"));
        CHECK(archive.read("xl/workbook.xml").find("fullCalcOnLoad") != std::string::npos);

        const auto back = sc::office::excel::read(file);
        CHECK_EQ(back.size(), std::size_t{2});
        CHECK_EQ(back.front().name(), std::string{"Prices"});
        CHECK(back["Empty"].empty());
        const auto &got = back["Prices"];
        CHECK_EQ(got.size(), std::size_t{5}); // rows 1, 2, 4, 5 and 10
        for (const auto &row: prices)
            for (const auto &cell: row) {
                if (cell.empty()) {
                    CHECK_MSG(!got.contains({cell.row, cell.column}), cell.name());
                    continue;
                }
                CHECK_MSG(got.contains({cell.row, cell.column}), cell.name());
                if (!got.contains({cell.row, cell.column})) continue;
                CHECK_MSG(got.at({cell.row, cell.column}).value == cell.value, cell.name());
                CHECK_MSG(got.at({cell.row, cell.column}).formula == cell.formula, cell.name());
            }
        CHECK(text(got, "C4") == "line\nbreak, bell\a, _x0041_ as it is, ünïcödé");
        CHECK_EQ(number(got, "A4"), 0.1);
        CHECK(std::get<bool>(got.at("D2").value) == false);
        CHECK(text(got, "D5") == "yes");
    }

    SECTION("Written the way Excel writes it");
    {
        const auto file = root / "excel.xlsx";
        {
            auto out = sc::archive::create(file);
            out.add("xl/workbook.xml", workbook_xml);
            out.add("xl/_rels/workbook.xml.rels", workbook_rels);
            out.add("xl/sharedStrings.xml", shared_strings);
            out.add("xl/worksheets/sheet1.xml", sheet1);
            out.add("xl/worksheets/sheet2.xml", sheet2);
            out.close();
        }
        const auto book = sc::office::excel::read(file);
        CHECK_EQ(book.size(), std::size_t{2});
        CHECK_EQ(book.front().name(), std::string{"Summary"}); // the workbook's order
        CHECK_EQ(number(book["Summary"], "B2"), -0.03);
        CHECK(book.contains("data & more"));
        if (book.contains("data & more")) {
            const auto &data = book["data & more"];
            CHECK(text(data, "A1") == "Item");
            CHECK(!data.contains("B1"));
            CHECK(text(data, "C1") == "Bold and plain");
            CHECK(!data.contains("A2"));
            CHECK_EQ(number(data, "A3"), 12.5);
            CHECK(std::get<bool>(data.at("B3").value));
            CHECK(text(data, "C3") == "日本"); // the reading left out
            // A row and its cells without r follow the ones before; formulas with their values.
            CHECK_EQ(number(data, "A4"), 25.0);
            CHECK_EQ(data.at("A4").formula, std::string{"=A3*2"});
            CHECK(text(data, "B4") == "xy");
            CHECK_EQ(data.at("B4").formula, std::string{R"(="x"&"y")"});
            CHECK(text(data, "C4") == "#DIV/0!");
            CHECK(text(data, "B6") == "inline");
            CHECK(!data.contains("C6")); // styled but empty
            CHECK(text(data, "D6") == "Tab\there");
            // A formula kept once for A7:A9: each cell's own, relative references moved.
            CHECK_EQ(data.at("A7").formula, std::string{R"(=A3+$B$1*C3&"A1"+'My sheet'!B2+SUM(D3:D4)+LOG10(1E5))"});
            CHECK_EQ(data.at("A8").formula, std::string{R"(=A4+$B$1*C4&"A1"+'My sheet'!B3+SUM(D4:D5)+LOG10(1E5))"});
            CHECK_EQ(data.at("A9").formula, std::string{R"(=A5+$B$1*C5&"A1"+'My sheet'!B4+SUM(D5:D6)+LOG10(1E5))"});
            CHECK_EQ(number(data, "A8"), 2.0);
            CHECK(is_empty(data.at("A9").value)); // never calculated
            CHECK_EQ(data.at("B8").formula, std::string{"=NOW()"});
        }
        // The same from memory.
        std::ifstream in{file, std::ios::binary};
        const auto from_memory = sc::office::excel::read_from(std::string{std::istreambuf_iterator<char>{in}, {}});
        CHECK_EQ(from_memory.size(), std::size_t{2});
        CHECK_EQ(from_memory.at("Data & more").size(), book.at("Data & more").size());
    }

    SECTION("Not a workbook, and what can't be written");
    {
        const auto zip = root / "other.zip";
        {
            auto out = sc::archive::create(zip);
            out.add("readme.txt", "hello");
            out.close();
        }
        CHECK_THROWS_AS((void) sc::office::excel::read(zip), std::runtime_error);
        CHECK_THROWS_AS((void) sc::office::excel::read_from("not a zip at all"), std::runtime_error);

        const auto file = root / "never.xlsx";
        CHECK_THROWS_AS(sc::office::excel::write({}, file), std::invalid_argument);
        sc::office::excel::workbook book;
        book["A"]["A1"] = std::numeric_limits<double>::quiet_NaN();
        CHECK_THROWS_AS(sc::office::excel::write(book, file), std::invalid_argument);
        CHECK(!fs::exists(file));
    }

    fs::remove_all(root);
    TEST_SUMMARY();
}
