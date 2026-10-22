# simply-cpp-office

`sc-office` reads Office documents. So far: the pictures in a Word document
(`.docx`); Excel (`.xlsx`) is to come. Office files are ZIP archives of XML
(Office Open XML), read with sc-compression and sc-web.

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

**TODO:** Excel: sheets, cells and their values; pictures in `.xlsx` and `.pptx`.
The old binary `.doc` and `.xls` formats aren't planned.

## Demo

`sc-office-demo FILE [--save FOLDER]` lists a `.docx`'s pictures in order, or
writes them into a folder as they're stored.

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
The test builds its own `.docx` (with sc-compression) and needs nothing else.
