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
prices["A1"].style = {.font = {.size = 14, .color = sc::color::White, .bold = true},
                      .background = sc::color{"#4472c4"}};
prices["B3"].style.font.bold = true;
prices["B3"].style.borders.top = {sc::office::excel::line::thin};
prices["B3"].style.borders.bottom = {sc::office::excel::line::doubled, sc::color{"#1f4e79"}};
prices["B3"].style.number_format = "#,##0.00";
prices["A1"].style.alignment = {.horizontal = sc::office::excel::horizontal::center, .wrap = true};
prices.width("A", 20);                   // in characters, as Excel shows it
sc::office::excel::write(out, "out.xlsx");
```

Workbooks, sheets and rows work like the standard containers (and nlohmann::json):

- **Cells** (`excel::cell`): `row` and `column` (from 0), `value` and `formula` (`"=C6"`, as Excel
  shows it; empty when there's none). A value (`excel::value`) is a `std::variant` of empty
  (`std::monostate`), a number (`double`), text (`std::string`, UTF-8) or `bool`. Dates are
  numbers, days since 1899-12-30 as Excel keeps them; an error such as `#DIV/0!` reads as its
  text. `cell = value` sets the value; `name()` gives `"B3"`. Its `style` (`excel::style`) is how
  it looks: `font.size` (an `sc::distance`, so a plain number is points), `font.color` and
  `background` (`sc::color`, a solid fill; alpha is left out), `font.bold`, `font.italic`,
  `font.underline` (`excel::underline`: `single`, `doubled`, `single_accounting`,
  `double_accounting`), and `borders`: `left`, `right`, `top` and `bottom`, each an `excel::edge`
  of a `line` (`excel::line`: `thin`, `medium`, `thick`, `doubled`, `dashed`, `dotted`, `hair` and
  Excel's dash-dot kinds) and a `color`; `excel::borders::all({excel::line::thin})` puts one edge
  on all four sides. `alignment`: `horizontal` (`excel::horizontal`: `general`, `left`, `center`,
  `right`, `fill`, `justify`, `center_across`, `distributed`), `vertical` (`excel::vertical`:
  `bottom`, `top`, `center`, `justify`, `distributed`), `wrap` and `indent` (0 to 250).
  `number_format` is Excel's format code (`"0.00"`, `"#,##0"`, `"0%"`, `"yyyy-mm-dd"`, `"@"`),
  empty for General; a date is a number shown as one by a date format. `name(line::doubled)` is
  Excel's `"double"`, `line_named("double")` the line (likewise for underlines and alignments).
  What isn't set is Excel's default (Calibri 11, automatic colour, no fill, no borders, General,
  aligned as Excel aligns text and numbers).
- **Addresses** (`excel::address`): `{row, column}`, or Excel's name, `"B3"`; `name()` gives it
  back, `column_name(26)` gives `"AA"` and `column_index("AA")` 26.
- **Rows** (`excel::row`): the cells that are set, in column order; `index()`, `contains(column)`,
  `at(column)`, `row[column]`, `erase(column)`, `size()`, `empty()`.
- **Sheets** (`excel::sheet`): `name()`; the rows that have cells, in order. `contains(address)`;
  `at(address)` throws `std::out_of_range` when there's no such cell; `sheet[address]` adds it
  (empty) when there's none, as `std::map` does (on a const sheet it's `at()`); `erase(address)`
  removes it, and its row with its last cell. `width(column)` is a column's width in characters
  (so many `0`s, as Excel shows it), none for Excel's default; `width(column, characters)` sets it
  (0 to 255; `std::nullopt` for the default), the column by number or letters (`"B"`); `widths()`
  lists them.
- **Workbooks** (`excel::workbook`): the sheets in order, `size()`, `empty()`, `front()`.
  `contains(name)`, `at(name)` and `book[name]` find a sheet ignoring case, as Excel does;
  `book[name]` adds it when there's none (on a const workbook it's `at()`). `add(name)` adds one,
  checking the name as Excel would: 1 to 31 characters, none of `: \ / ? * [ ]`, not starting or
  ending with `'`, not already used. A sheet stays where it is as more are added.
- **Reading** (`read(file)`, `read_from(bytes)`): shared and inline strings, rich text (its runs
  joined, phonetic readings left out), booleans, errors, formulas and the values they last
  calculated (none when they never were, as in files some libraries write). A formula Excel keeps
  once for a range is given to each of its cells, its relative references moved. Styles come with
  their theme and indexed colours turned into `sc::color`s (tints applied), what's the same as the
  workbook's default font left unset, and number formats as their codes (Excel's built-in ones
  too: 14, the reader's own short date, reads as `"mm-dd-yy"`); a cell with a style and no value
  is kept. Column widths
  are those set by hand (`customWidth`). A file that isn't an `.xlsx` throws `std::runtime_error`.
- **Writing** (`write(book, file)`, or `bytes(book)` for the file's bytes): each cell's value,
  formula and style, text inline, and the columns' widths; Excel calculates the formulas when it
  opens the file. It throws `std::invalid_argument` for a workbook without sheets, a number that
  isn't finite or a font size outside 1 to 409 points, and then leaves no file behind.

Not covered yet: fonts other than Calibri, strikethrough, diagonal borders, text rotation, row
heights, merged cells, pictures and charts.

**TODO:** pictures in `.xlsx` and `.pptx`. The old binary `.doc` and `.xls` formats aren't
planned.

## Demos

`sc-office-word-demo FILE [--save FOLDER]` lists a `.docx`'s pictures in order,
or writes them into a folder as they're stored.

<!-- sc-example: examples/sc-office-word-demo.cpp -->
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

`sc-office-excel-demo FILE` prints an `.xlsx`'s cells, sheet by sheet, with
their values and formulas; `sc-office-excel-demo --write FILE` writes a small
workbook, with a styled, centred header, widened columns, prices to two decimals, a ruled total
and an italic note.

<!-- sc-example: examples/sc-office-excel-demo.cpp -->
```cpp
for (const auto &sheet: sc::office::excel::read(file)) {
    std::cout << "== " << sheet.name() << '\n';
    for (const auto &row: sheet)
        for (const auto &cell: row)
            std::cout << cell.name() << '\t' << text_of(cell.value) << (cell.formula.empty() ? "" : "\t" + cell.formula) << '\n';
}
```
<!-- /sc-example -->

## Building and testing

Run `bash scripts/install.sh` to configure, build, test, package and install.
The tests build their own `.docx` and `.xlsx` files (with sc-archive) and need
nothing else; the Excel one also writes workbooks and reads them back.
