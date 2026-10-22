// Lists the pictures in a Word document, in the order they appear, or saves them with --save.
//
//     sc-office-demo statement.docx
//     sc-office-demo statement.docx --save pictures

#include <office.h>
#include <params.h>

#include <filesystem>
#include <fstream>
#include <iostream>

int main(int argc, char **argv) {
    const sc::params args{argc, argv, {
        {"name", "sc-office-demo"},
        {"description", "Lists the pictures in a Word document (.docx): those in its text in the order they appear, then "
                        "the rest; with --save, writes them into a folder as they're stored."},
        {"options", {
            {"save", {{"short", "s"}, {"value", "FOLDER"}, {"default", ""}, {"help", "write the pictures into FOLDER"}}}}},
        {"positional", {
            {"file", {{"help", "the .docx"}}}}}}};
    try {
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
    } catch (const std::exception &error) {
        std::cerr << "sc-office-demo: " << error.what() << '\n';
        return 1;
    }
    return 0;
}
