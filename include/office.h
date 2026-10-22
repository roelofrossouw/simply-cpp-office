#pragma once

#include <filesystem>
#include <string>
#include <string_view>
#include <vector>

// Office documents: Word (.docx) so far; Excel (.xlsx) to come. They're ZIP archives of XML
// (Office Open XML), read with sc-archive and sc-web.
//
//     for (const auto &picture: sc::office::word::images("statement.docx"))
//         if (picture.raster()) process(picture.name, picture.bytes);
//
// TODO: Excel: sheets, cells and their values; pictures in .xlsx and .pptx.
namespace sc::office {
    // A picture embedded in a document, as stored: name is its file name in the document
    // ("image1.png"), bytes its file's bytes.
    struct picture {
        std::string name;
        std::string bytes;

        // Whether it's an image sc::image reads (JPEG, PNG, GIF, BMP, TIFF, WebP), rather than a
        // drawing (EMF, WMF, SVG), by its name.
        [[nodiscard]] bool raster() const;
    };

    namespace word {
        // The pictures in a Word document: those in its text first, in the order they appear (each
        // once), then any others it holds (in headers and footers, or not placed). Throws
        // std::runtime_error for a file that isn't a .docx (not a ZIP archive, or no
        // word/document.xml).
        std::vector<picture> images(const std::filesystem::path &docx);
        // The same, for a .docx held in memory.
        std::vector<picture> images_in(std::string docx);
    }
}
