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

    // Styles as Excel writes them: theme colours with tints, indexed and rgb ones, a default font
    // of 11 points in the theme's dark 1.
    const std::string styled_rels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>)";
    const std::string styled_workbook = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Styled" sheetId="1" r:id="rId1"/></sheets>
</workbook>)";
    const std::string theme = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office">
  <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
  <a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
  <a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
  <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
  <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
  <a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme></a:themeElements></a:theme>)";
    const std::string styles = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="4">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>
    <font><b/><sz val="16"/><color theme="0"/><name val="Calibri"/></font>
    <font><b val="0"/><sz val="11"/><color rgb="FFC00000"/><name val="Calibri"/></font>
    <font><sz val="11"/><color indexed="10"/><name val="Calibri"/></font>
    <font><i/><u val="double"/><sz val="11"/><color theme="1"/><name val="Calibri"/></font>
    <font><u/><sz val="11"/><color theme="1"/><name val="Calibri"/></font>
  </fonts>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><start style="thin"><color indexed="64"/></start><end style="mediumDashDot"/><top style="double"><color theme="4"/></top><bottom style="madeUp"/><diagonal/></border>
  </borders>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor theme="4" tint="0.79998168889431442"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor theme="4"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
</styleSheet>)";
    const std::string styled_sheet = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="1" width="20.7109375" customWidth="1"/>
    <col min="2" max="4" width="9.140625" customWidth="1"/>
    <col min="5" max="16384" width="9.140625" style="1"/>
  </cols>
  <sheetData>
    <row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Header</t></is></c><c r="B1" s="2"><v>1</v></c><c r="C1" s="3"/><c r="D1" s="4"/><c r="E1" s="9"><v>2</v></c><c r="F1" s="5"/><c r="G1" s="6"><v>3</v></c></row>
  </sheetData>
</worksheet>)";

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
        // Erasing, as std::map does: a row goes with its last cell.
        CHECK_EQ(sheet.erase("C2"), std::size_t{1});
        CHECK_EQ(sheet.erase("C2"), std::size_t{0});
        CHECK_EQ(sheet.erase("Z9"), std::size_t{0});
        CHECK_EQ(sheet.size(), std::size_t{1});
        CHECK_EQ(sheet.erase("A4"), std::size_t{1});
        CHECK_EQ(sheet.size(), std::size_t{1}); // E4 is still there
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
        // As bytes: the same workbook.
        const auto bytes = sc::office::excel::bytes(book);
        CHECK(bytes.starts_with("PK"));
        CHECK(text(sc::office::excel::read_from(bytes)["Prices"], "A2") == "Bread");
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

    SECTION("Styles and column widths, written and read back");
    {
        using sc::office::excel::column_index;
        CHECK_EQ(column_index("A"), std::size_t{0});
        CHECK_EQ(column_index("aa"), std::size_t{26});
        CHECK_EQ(column_index("XFD"), std::size_t{16383});
        for (const auto *wrong: {"", "A1", "XFE", "-"}) CHECK_THROWS_AS((void) column_index(wrong), std::invalid_argument);

        sc::office::excel::workbook book;
        auto &sheet = book["Styled"];
        sheet["A1"] = "Header";
        sheet["A1"].style = {.font = {.size = 14, .color = sc::color{"#1f4e79"}, .bold = true}, .background = sc::color{"#ddebf7"}};
        sheet["B1"] = "Bold";
        sheet["B1"].style.font.bold = true;
        sheet["C1"].style.background = sc::color::Red; // a style, no value: kept
        sheet["A2"] = 1.0;
        sheet["A2"].style = sheet["A1"].style;          // the same style: written once
        sheet["B2"] = 2.0;
        sheet["B2"] = 3.0;                              // setting the value keeps the style
        using sc::office::excel::line;
        using sc::office::excel::underline;
        sheet["D2"] = "Fancy";
        sheet["D2"].style.font.italic = true;
        sheet["D2"].style.font.underline = underline::double_accounting;
        sheet["D2"].style.borders = sc::office::excel::borders::all({line::thin});
        sheet["D2"].style.borders.bottom = {line::doubled, sc::color{"#1f4e79"}};
        sheet["E2"].style.borders = sheet["D2"].style.borders; // the same borders: written once
        sheet.width("A", 30);
        sheet.width(1, 12.5);
        sheet.width("C", 12.5);                         // B and C: one <col>
        sheet.width("E", 4);
        CHECK(sheet.width("A") == 30.0);
        CHECK(!sheet.width("D"));
        sheet.width("E", std::nullopt);
        CHECK(!sheet.width(4));
        CHECK_EQ(sheet.widths().size(), std::size_t{3});
        CHECK_THROWS_AS(sheet.width("A", -1), std::out_of_range);
        CHECK_THROWS_AS(sheet.width("A", 256), std::out_of_range);
        CHECK_THROWS_AS(sheet.width("A", std::numeric_limits<double>::quiet_NaN()), std::out_of_range);
        CHECK_THROWS_AS(sheet.width(sc::office::excel::max_columns, 1), std::out_of_range);
        CHECK(sheet.width("A") == 30.0);

        const auto bytes = sc::office::excel::bytes(book);
        const auto archive = sc::archive::from_memory(bytes);
        const auto styles_xml = archive.read("xl/styles.xml");
        CHECK(styles_xml.find(R"(<cellXfs count="6">)") != std::string::npos); // plain, A1 and A2, B1, C1, D2, E2
        CHECK(styles_xml.find(R"(<borders count="2">)") != std::string::npos);
        CHECK(styles_xml.find(R"(<i/><u val="doubleAccounting"/>)") != std::string::npos);
        const auto xml = archive.read("xl/worksheets/sheet1.xml");
        CHECK(xml.find(R"(<col min="2" max="3" width="12.5" customWidth="1"/>)") != std::string::npos);
        CHECK(xml.find(R"(<c r="B2"><v>3</v></c>)") != std::string::npos); // plain: no s=

        const auto back = sc::office::excel::read_from(bytes);
        const auto &got = back["Styled"];
        const auto &header = got.at("A1").style;
        CHECK(header.font.bold);
        CHECK(header.font.size && header.font.size->points() == 14);
        CHECK(header.font.color && header.font.color->to_hex() == "#1f4e79");
        CHECK(header.background && header.background->to_hex() == "#ddebf7");
        CHECK(got.at("A2").style == header);
        CHECK(got.at("B1").style.font.bold && !got.at("B1").style.font.size && !got.at("B1").style.font.color);
        CHECK(got.contains("C1") && got.at("C1").style.background && got.at("C1").style.background->to_hex() == "#ff0000");
        CHECK(got.at("B2").style.empty());
        CHECK(got.widths() == sheet.widths());
        const auto &fancy = got.at("D2").style;
        CHECK(fancy.font.italic && fancy.font.underline == underline::double_accounting);
        CHECK(fancy.borders.left == (sc::office::excel::edge{line::thin}));
        CHECK(fancy.borders.top.line == line::thin && fancy.borders.right.line == line::thin);
        CHECK(fancy.borders.bottom.line == line::doubled && fancy.borders.bottom.color && fancy.borders.bottom.color->to_hex() == "#1f4e79");
        CHECK(got.contains("E2") && got.at("E2").style.borders.left.line == line::thin);

        CHECK(sc::office::excel::name(line::medium_dash_dot_dot) == "mediumDashDotDot");
        CHECK(sc::office::excel::line_named("double") == line::doubled);
        CHECK(sc::office::excel::underline_named("singleAccounting") == underline::single_accounting);
        CHECK_THROWS_AS((void) sc::office::excel::line_named("Thin"), std::invalid_argument);
        CHECK_THROWS_AS((void) sc::office::excel::underline_named("wavy"), std::invalid_argument);

        sheet["D1"].style.font.size = 500;
        CHECK_THROWS_AS((void) sc::office::excel::bytes(book), std::invalid_argument);
    }

    SECTION("Styles and column widths the way Excel writes them");
    {
        const auto file = root / "styled.xlsx";
        {
            auto out = sc::archive::create(file);
            out.add("xl/workbook.xml", styled_workbook);
            out.add("xl/_rels/workbook.xml.rels", styled_rels);
            out.add("xl/theme/theme1.xml", theme);
            out.add("xl/styles.xml", styles);
            out.add("xl/worksheets/sheet1.xml", styled_sheet);
            out.close();
        }
        const auto book = sc::office::excel::read(file);
        const auto &sheet = book["Styled"];
        // Bold, 16 points, the theme's light 1 (white) on accent 1.
        const auto &header = sheet.at("A1").style;
        CHECK(header.font.bold);
        CHECK(header.font.size && header.font.size->points() == 16);
        CHECK(header.font.color && header.font.color->to_hex() == "#ffffff");
        CHECK(header.background && header.background->to_hex() == "#4472c4");
        // Red text, not bold (b val="0"), the default size left unset.
        const auto &red = sheet.at("B1").style;
        CHECK(!red.font.bold && !red.font.size);
        CHECK(red.font.color && red.font.color->to_hex() == "#c00000");
        // Indexed 10 (red) on accent 1 lightened 80%, as Excel shows it: #dae3f3 (or a step off).
        CHECK(sheet.contains("C1"));
        if (sheet.contains("C1")) {
            const auto &tinted = sheet.at("C1").style;
            CHECK(tinted.font.color && tinted.font.color->to_hex() == "#ff0000");
            CHECK(tinted.background);
            if (tinted.background) {
                CHECK(std::abs(tinted.background->red() * 255 - 0xDA) < 1.5);
                CHECK(std::abs(tinted.background->green() * 255 - 0xE3) < 1.5);
                CHECK(std::abs(tinted.background->blue() * 255 - 0xF3) < 1.5);
            }
        }
        CHECK(!sheet.contains("D1"));                 // a date format only: nothing we keep
        // Italic, double underline; borders by start and end (left and right), theme colours, a
        // line Excel doesn't have drawn thin, indexed 64 (automatic) as no colour.
        CHECK(sheet.contains("F1"));
        if (sheet.contains("F1")) {
            const auto &look = sheet.at("F1").style;
            CHECK(look.font.italic && !look.font.bold);
            CHECK(look.font.underline == sc::office::excel::underline::doubled);
            CHECK(look.borders.left.line == sc::office::excel::line::thin && !look.borders.left.color);
            CHECK(look.borders.right.line == sc::office::excel::line::medium_dash_dot);
            CHECK(look.borders.top.line == sc::office::excel::line::doubled);
            CHECK(look.borders.top.color && look.borders.top.color->to_hex() == "#4472c4");
            CHECK(look.borders.bottom.line == sc::office::excel::line::thin);
        }
        CHECK(sheet.at("G1").style.font.underline == sc::office::excel::underline::single); // <u/>: single
        CHECK(sheet.at("E1").style.empty());          // a style that isn't there
        CHECK_EQ(sheet.widths().size(), std::size_t{4}); // A, and B to D; not the styled rest
        CHECK(sheet.width("A") == 20.7109375);
        CHECK(sheet.width("D") == 9.140625);
        CHECK(!sheet.width("E"));
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
