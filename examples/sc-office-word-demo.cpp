// A Word document's text, its list numbers as Word shows them; or, with --pictures, the pictures in
// it in the order they appear, saved into a folder with --save.
//
//     sc-office-word-demo agreement.docx
//     sc-office-word-demo agreement.docx --headers --notes
//     sc-office-word-demo statement.docx --pictures --save pictures

#include <sc/office/office.h>
#include <sc/core/params.h>

#include <filesystem>
#include <fstream>
#include <iostream>

int main(int argc, char **argv) {
    const sc::params args{argc, argv, {
        {"name", "sc-office-word-demo"},
        {"description", "Prints a Word document's (.docx) text: a line per paragraph with its list number as Word shows it "
                        "(\"4.1.2 ...\"), a table a row per line with its cells separated by \" | \". With --pictures, lists "
                        "its pictures instead: those in its text in the order they appear, then the rest; with --save, "
                        "writes them into a folder as they're stored."},
        {"options", {
            {"headers", {{"type", "boolean"}, {"help", "the headers' and footers' text too, after the body"}}},
            {"notes", {{"type", "boolean"}, {"help", "the footnotes' and endnotes' text too, at the end"}}},
            {"pictures", {{"short", "p"}, {"type", "boolean"}, {"help", "list the pictures instead of the text"}}},
            {"save", {{"short", "s"}, {"value", "FOLDER"}, {"default", ""}, {"help", "with --pictures: write them into FOLDER"}}}}},
        {"positional", {
            {"file", {{"help", "the .docx"}}}}}}};
    try {
        // [readme]
        const std::string file = args.as<std::string>("file");
        if (!args.as<bool>("pictures")) {
            std::cout << sc::office::word::text(file, {.headers_footers = args.as<bool>("headers"), .notes = args.as<bool>("notes")}) << '\n';
            return 0;
        }
        const auto pictures = sc::office::word::images(file);
        const std::filesystem::path folder = args.as<std::string>("save");
        if (!folder.empty()) std::filesystem::create_directories(folder);
        for (const auto &picture: pictures) {
            std::cout << picture.name << "  " << picture.bytes.size() << " bytes" << (picture.raster() ? "" : "  (a drawing)") << '\n';
            if (!folder.empty()) std::ofstream{folder / picture.name, std::ios::binary} << picture.bytes;
        }
        std::cout << pictures.size() << " pictures\n";
        // [/readme]
    } catch (const std::exception &error) {
        std::cerr << "sc-office-word-demo: " << error.what() << '\n';
        return 1;
    }
    return 0;
}
