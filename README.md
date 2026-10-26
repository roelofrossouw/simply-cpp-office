# simply-cpp-office

`sc-office` reads Office documents: the pictures in a Word document (`.docx`),
and an Excel workbook's cells (`.xlsx`), which it writes too. Office files are
ZIP archives of XML (Office Open XML), read and written with sc-archive and
sc-web.

**Documentation:** the [simply-cpp wiki](https://github.com/roelofrossouw/simply-cpp/wiki) has an [Office guide](https://github.com/roelofrossouw/simply-cpp/wiki/Office) and the [sc-office reference](https://github.com/roelofrossouw/simply-cpp/wiki/Reference-sc-office).

```cpp
#include <office.h>

for (const auto &picture: sc::office::word::images("statement.docx")) {
    if (!picture.raster()) continue;     // EMF, WMF, SVG: drawings, not images
    picture.name;                        // "image1.png"
    picture.bytes;                       // the file as stored: no re-encoding, full resolution
}
```

`word::images()` gives the pictures in the document's text first, in the order
they appear, each once (DrawingML pictures and the older VML ones), then any
others the document holds: those in headers and footers, or not placed. Linked
pictures (outside the file) aren't included. `word::images_in(bytes)` reads a
`.docx` held in memory. A file that isn't a `.docx` throws `std::runtime_error`.

## Excel

```cpp
#include <excel.h>                       // or office.h, which includes it

// Reading: the sheets in order, the rows that have cells, the cells that are set.
const auto book = sc::office::excel::read("prices.xlsx");
for (const auto &sheet: book)
    for (const auto &row: sheet)
        for (const auto &cell: row)
            cell.row, cell.column, cell.value, cell.formula;   // 3, 4, "Hello", "=C6"
if (book.contains("Prices") && book["Prices"].contains("B2"))
    price = std::get<double>(book["Prices"].at("B2").value);

// Writing.
sc::office::excel::workbook out;
auto &prices = out["Prices"];            // added, as std::map would
prices["A1"] = "Item";
prices[{1, 1}] = 18.5;                   // B2: {row, column}, counted from 0
prices["B3"].formula = "=B2*2";
sc::office::excel::write(out, "out.xlsx");
```

Workbooks, sheets and rows work like the standard containers (and nlohmann::json):

- **Cells** (`excel::cell`): `row` and `column` (from 0), `value` and `formula` (`"=C6"`, as Excel
  shows it; empty when there's none). A value (`excel::value`) is a `std::variant` of empty
  (`std::monostate`), a number (`double`), text (`std::string`, UTF-8) or `bool`. Dates are
  numbers, days since 1899-12-30 as Excel keeps them; an error such as `#DIV/0!` reads as its
  text. `cell = value` sets the value; `name()` gives `"B3"`.
- **Addresses** (`excel::address`): `{row, column}`, or Excel's name, `"B3"`; `name()` gives it
  back, `column_name(26)` gives `"AA"`.
- **Rows** (`excel::row`): the cells that are set, in column order; `index()`, `contains(column)`,
  `at(column)`, `row[column]`, `size()`, `empty()`.
- **Sheets** (`excel::sheet`): `name()`; the rows that have cells, in order. `contains(address)`;
  `at(address)` throws `std::out_of_range` when there's no such cell; `sheet[address]` adds it
  (empty) when there's none, as `std::map` does (on a const sheet it's `at()`).
- **Workbooks** (`excel::workbook`): the sheets in order, `size()`, `empty()`, `front()`.
  `contains(name)`, `at(name)` and `book[name]` find a sheet ignoring case, as Excel does;
  `book[name]` adds it when there's none (on a const workbook it's `at()`). `add(name)` adds one,
  checking the name as Excel would: 1 to 31 characters, none of `: \ / ? * [ ]`, not starting or
  ending with `'`, not already used. A sheet stays where it is as more are added.
- **Reading** (`read(file)`, `read_from(bytes)`): shared and inline strings, rich text (its runs
  joined, phonetic readings left out), booleans, errors, formulas and the values they last
  calculated (none when they never were, as in files some libraries write). A formula Excel keeps
  once for a range is given to each of its cells, its relative references moved. A file that
  isn't an `.xlsx` throws `std::runtime_error`.
- **Writing** (`write(book, file)`): each cell's value and formula, text inline, no styles; Excel
  calculates the formulas when it opens the file. It throws `std::invalid_argument` for a
  workbook without sheets or a number that isn't finite, and then leaves no file behind.

Not covered: formatting, styles, number formats, column widths, merged cells, pictures and
charts.

**TODO:** pictures in `.xlsx` and `.pptx`. The old binary `.doc` and `.xls` formats aren't
planned.

## Demo

`sc-office-demo FILE [--save FOLDER]` lists a `.docx`'s pictures in order, or
writes them into a folder as they're stored. For an `.xlsx` it prints each
sheet's cells with their values and formulas; `sc-office-demo --write FILE`
writes a small workbook.

<!-- sc-example: examples/sc-office-demo.cpp -->
```cpp
const auto pictures = sc::office::word::images(args.as<std::string>("file"));
const std::filesystem::path folder = args.as<std::string>("save");
if (!folder.empty()) std::filesystem::create_directories(folder);
for (const auto &picture: pictures) {
    std::cout << picture.name << "  " << picture.bytes.size() << " bytes" << (picture.raster() ? "" : "  (a drawing)") << '\n';
    if (!folder.empty()) std::ofstream{folder / picture.name, std::ios::binary} << picture.bytes;
}
std::cout << pictures.size() << " pictures\n";
```
<!-- /sc-example -->

## Building and testing

Run `bash scripts/install.sh` to configure, build, test, package and install.
The tests build their own `.docx` and `.xlsx` files (with sc-archive) and need
nothing else; the Excel one also writes workbooks and reads them back.
