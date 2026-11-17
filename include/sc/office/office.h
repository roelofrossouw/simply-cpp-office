#pragma once

#include <sc/office/excel.h>

#include <filesystem>
#include <string>
#include <string_view>
#include <vector>

// Office documents: a Word document's text and pictures (.docx) here, an Excel workbook's cells
// (.xlsx) in excel.h, which this includes. They're ZIP archives of XML (Office Open XML), read with
// sc-archive and sc-web.
//
//     const auto agreement = sc::office::word::text("agreement.docx"); // "4.1.2 the Purchaser ..."
//     for (const auto &picture: sc::office::word::images("statement.docx"))
//         if (picture.raster()) process(picture.name, picture.bytes);
//
// TODO: pictures in .xlsx and .pptx.
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
        // What text() reads besides the body.
        struct text_options {
            bool headers_footers = false; // the headers', then the footers' text, after the body
            bool notes = false;           // the footnotes' and endnotes' text, after that
        };

        // A Word document's text, as a reader (a person or a language model) would take it: a line
        // per paragraph, with its list number in front as Word shows it ("4.1.2 the Purchaser ...":
        // numbering.xml's levels and formats, set on the paragraph or through its style, counted per
        // list, a level restarting those below it; bullets get none); tabs as \t, line breaks as \n;
        // a table a row per line, its cells separated by " | "; content controls and text boxes as
        // part of the text. Deleted text (tracked changes), field codes and the copies kept for older
        // readers (mc:Fallback) are left out. Spacing tidied: at most one blank line between
        // paragraphs. Throws std::runtime_error for a file that isn't a .docx (an old .doc too).
        std::string text(const std::filesystem::path &docx, const text_options &options = {});
        // The same, for a .docx held in memory.
        std::string text_in(std::string docx, const text_options &options = {});

        // The pictures in a Word document: those in its text first, in the order they appear (each
        // once), then any others it holds (in headers and footers, or not placed). Throws
        // std::runtime_error for a file that isn't a .docx (not a ZIP archive, or no
        // word/document.xml).
        std::vector<picture> images(const std::filesystem::path &docx);
        // The same, for a .docx held in memory.
        std::vector<picture> images_in(std::string docx);
    }
}
