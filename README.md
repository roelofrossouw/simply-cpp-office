# simply-cpp-office

`sc-office` reads Office documents: the pictures in a Word document (`.docx`),
and an Excel workbook's cells (`.xlsx`), which it writes too, as it reads and
writes CSV files with the same access to their values. Office files are
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
prices.height(0, 24);                    // row 1, in points
prices.merge("A5:B5");                   // shown as one cell, A5
prices.hidden_column("C", true);         // a hidden column
prices.hidden_row(9, true);              // a hidden row (row 10)
prices.freeze("A2");                     // row 1 stays in view as the rest scrolls
sc::office::excel::write(out, "out.xlsx");
```

Workbooks, sheets and rows work like the standard containers (and nlohmann::json):

- **Cells** (`excel::cell`): `row` and `column` (from 0), `value` and `formula` (`"=C6"`, as Excel
  shows it; empty when there's none). A value (`excel::value`) is a `std::variant` of empty
  (`std::monostate`), a number (`double`), text (`std::string`, UTF-8) or `bool`. Dates are
  numbers, days since 1899-12-30 as Excel keeps them; an error such as `#DIV/0!` reads as its
  text. `cell = value` sets the value; `name()` gives `"B3"`. Its `style` (`excel::style`) is how
  it looks: `font.size` (an `sc::distance`, so a plain number is points), `font.color` and
  `background` (`sc::color`, a solid fill; alpha is left out), `font.bold`, `font.italic`, `font.strike`,
  `font.underline` (`excel::underline`: `single`, `doubled`, `single_accounting`,
  `double_accounting`), and `borders`: `left`, `right`, `top` and `bottom`, each an `excel::edge`
  of a `line` (`excel::line`: `thin`, `medium`, `thick`, `doubled`, `dashed`, `dotted`, `hair` and
  Excel's dash-dot kinds) and a `color`; `excel::borders::all({excel::line::thin})` puts one edge
  on all four sides; `diagonal_up` and `diagonal_down` are the diagonals (Excel draws both alike,
  so two different ones throw on writing). `alignment`: `horizontal` (`excel::horizontal`:
  `general`, `left`, `center`, `right`, `fill`, `justify`, `center_across`, `distributed`),
  `vertical` (`excel::vertical`:
  `bottom`, `top`, `center`, `justify`, `distributed`), `wrap`, `indent` (0 to 250),
  `rotation` (degrees counter-clockwise, -90 to 90) and `stacked` (letters one under another).
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
  lists them. `height(row)` is a row's height in points (an `sc::distance`), none for Excel's
  default; `height(row, points)` sets it (0 to 409, 0 hides it; `std::nullopt` for the default)
  and `heights()` lists them; a row with a height and no cells is kept. `merge(range)` merges a
  block of cells into one, shown as its top left (throwing for one cell or an overlap with a merge
  already there), `unmerge(range)` undoes exactly that merge, `merged()` lists them.
  `hidden_column(column, true)` hides a column (by number or letters; `false` shows it again,
  keeping its width), `hidden_column(column)` says whether it is, `hidden_columns()` lists them;
  `hidden_row(row, true)`, `hidden_row(row)` and `hidden_rows()` do the same for rows (a hidden row
  with no cells is kept). `freeze(address)` freezes the
  rows above that cell and the columns left of it, as Excel's Freeze Panes does there
  (`freeze("B2")` keeps row 1 and column A in view; `freeze("A1")` freezes nothing);
  `unfreeze()`, and `frozen()` gives the cell, or none.
- **Ranges** (`excel::range`): `"A1:D3"`, or two opposite corners (`{"D3", "A1"}` is A1:D3);
  `first` and `last` (top left, bottom right), `name()`, `contains(address)`, `overlaps(range)`,
  `size()`.
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
  are those set by hand (`customWidth`), and row heights likewise (`customHeight`); merged
  cells too, one overlapping an earlier one left out; hidden columns and rows, and frozen panes (by how
  many rows and columns they keep, whatever the file was scrolled to). A file that isn't an `.xlsx` throws `std::runtime_error`.
- **Writing** (`write(book, file)`, or `bytes(book)` for the file's bytes): each cell's value,
  formula and style, text inline, and the columns' widths; Excel calculates the formulas when it
  opens the file. It throws `std::invalid_argument` for a workbook without sheets, a number that
  isn't finite or a font size outside 1 to 409 points, and then leaves no file behind.

Not covered yet: fonts other than Calibri, split (unfrozen) panes, pictures and charts.

**TODO:** pictures in `.xlsx` and `.pptx`. The old binary `.doc` and `.xls` formats aren't
planned.

## CSV

`csv.h` reads and writes CSV files (RFC 4180) as a sheet of values with the same
access as an Excel sheet, so code that only uses data works on both: cells by
name (`"B3"`) or place (`{2, 1}`), `at()`, `contains()`, `operator[]` adding a
cell as `std::map` does, `erase()`, and rows and cells in order. The values are
Excel's (`std::monostate`, `double`, `std::string`, `bool`); there are no styles,
widths or formulas.

```cpp
#include <csv.h>

const auto prices = sc::office::csv::read("prices.csv");   // , or ; found by itself
for (const auto &row: prices)
    for (const auto &cell: row)
        cell.row, cell.column, cell.value;                  // 3, 4, 12.5
if (prices.contains("B2")) total = std::get<double>(prices.at("B2").value);

sc::office::csv::sheet out;
out["A1"] = "Item";
out[{1, 1}] = 12.5;                                         // B2
sc::office::csv::write(out, "out.csv");                     // or bytes(out)

sc::office::csv::write(book["Prices"], "prices.csv");       // an Excel sheet's values
const auto xlsx_sheet = prices.to_excel();                  // and the other way
```

- **Quoting**: a field with the delimiter, a quote or a line break is quoted and a
  quote in it doubled (`"He said ""hi"""`); a quoted field may hold line breaks.
  Lines end in CRLF, LF or CR when read, and CRLF when written. A UTF-8 byte order
  mark is skipped when reading, and written with `.byte_order_mark = true`
  (Excel needs it to read UTF-8 text as such).
- **The delimiter**: `{.delimiter = ';'}` gives one; by default it's found: a comma,
  unless the file reads more evenly with semicolons (as Excel writes CSV where
  the decimal separator is a comma). `find_delimiter(text)` says which.
- **Values**: plain numbers (`12.5`, `-3`, `1e6`, but not `007`, `+1` or `.5`) read
  as numbers and TRUE or FALSE (any case) as true/false, as Excel reads them; a
  quoted field stays text. `{.typed = false}` reads every field as text. Writing
  quotes text that would read back as a number or TRUE/FALSE, so it stays text.
- **Empty fields** are cells that aren't set, so rows keep their places; a sheet is
  written from A1 to its last row and column.
- `read(file)` names the sheet after the file (`prices`); a quoted field that isn't
  closed throws `std::runtime_error`.

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
workbook, with a styled, centred, taller header kept in view, widened columns, a hidden cost
column, prices to two decimals, a ruled total, an italic note merged across two columns, a
discontinued item struck through with its price crossed out, a hidden row and a label turned 45
degrees.

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
nothing else; the Excel one also writes workbooks and reads them back, and the
CSV one checks its quoting, delimiters and values, and the same code on a CSV
sheet and an Excel sheet.
