// Pictures out of a .docx built here: the text's order (DrawingML and the older VML), each once,
// then the ones not placed; external links skipped; raster() by name; not-a-docx errors.
#include <archive.h>
#include <office.h>
#include <sc_test.h>

#include <filesystem>
#include <fstream>
#include <random>
#include <sstream>
#include <string>

namespace fs = std::filesystem;

namespace {
    const std::string body = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body>
    <w:p><w:r><w:t>Front</w:t></w:r></w:p>
    <w:p><w:r><w:drawing><a:graphic><a:graphicData><a:blip r:embed="rId7"/></a:graphicData></a:graphic></w:drawing></w:r></w:p>
    <w:p><w:r><w:t>Back</w:t></w:r></w:p>
    <w:p><w:r><w:drawing><a:graphic><a:graphicData><a:blip r:embed="rId5"/></a:graphicData></a:graphic></w:drawing></w:r></w:p>
    <w:p><w:r><w:pict><v:shape><v:imagedata r:id="rId9"/></v:shape></w:pict></w:r></w:p>
    <w:p><w:r><w:drawing><a:graphic><a:graphicData><a:blip r:embed="rId7"/></a:graphicData></a:graphic></w:drawing></w:r></w:p>
    <w:p><w:r><w:drawing><a:graphic><a:graphicData><a:blip r:embed="rId11"/></a:graphicData></a:graphic></w:drawing></w:r></w:p>
  </w:body>
</w:document>)";
    const std::string rels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.jpeg"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/word/media/image3.gif"/>
  <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.com/x.png" TargetMode="External"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>)";

    std::string slurp(const fs::path &file) {
        std::ifstream in{file, std::ios::binary};
        std::stringstream all;
        all << in.rdbuf();
        return all.str();
    }
}

int main() {
    using sc::office::picture;
    using std::string;
    const auto root = fs::temp_directory_path() / ("sc-office-test-" + std::to_string(std::random_device{}()));
    fs::create_directories(root);
    const auto docx = root / "test.docx";
    {
        auto out = sc::archive::create(docx);
        out.add("[Content_Types].xml", "<Types/>");
        out.add("word/document.xml", body);
        out.add("word/_rels/document.xml.rels", rels);
        out.add("word/media/image1.jpeg", "JPEG bytes", sc::archive::method::store);
        out.add("word/media/image2.png", "PNG bytes", sc::archive::method::store);
        out.add("word/media/image3.gif", "GIF bytes");
        out.add("word/media/image4.emf", "EMF bytes"); // a header's, not in the text
        out.close();
    }

    SECTION("Pictures: the text's order, each once, then the rest");
    {
        const auto pictures = sc::office::word::images(docx);
        CHECK_EQ(pictures.size(), std::size_t{4});
        CHECK_EQ(pictures[0].name, string{"image2.png"});  // placed first, though named second
        CHECK_EQ(pictures[0].bytes, string{"PNG bytes"});
        CHECK_EQ(pictures[1].name, string{"image1.jpeg"});
        CHECK_EQ(pictures[2].name, string{"image3.gif"});  // VML, an absolute target
        CHECK_EQ(pictures[3].name, string{"image4.emf"});  // not placed: last
        CHECK(pictures[0].raster() && pictures[1].raster() && pictures[2].raster());
        CHECK(!pictures[3].raster());
    }

    SECTION("From memory: the same");
    {
        const auto pictures = sc::office::word::images_in(slurp(docx));
        CHECK_EQ(pictures.size(), std::size_t{4});
        CHECK_EQ(pictures[1].bytes, string{"JPEG bytes"});
    }

    SECTION("raster() by the name's extension, any case");
    {
        CHECK((picture{"a.JPG", ""}.raster()));
        CHECK((picture{"scan.TIFF", ""}.raster()));
        CHECK((!picture{"logo.wmf", ""}.raster()));
        CHECK((!picture{"chart.svg", ""}.raster()));
    }

    SECTION("Not a .docx: runtime_error");
    {
        { std::ofstream{root / "plain.txt"} << "text"; }
        CHECK_THROWS_AS((void) sc::office::word::images(root / "plain.txt"), std::runtime_error);
        const auto other = root / "other.zip";
        {
            auto out = sc::archive::create(other);
            out.add("readme.txt", "no word/document.xml here");
            out.close();
        }
        CHECK_THROWS_AS((void) sc::office::word::images(other), std::runtime_error);
    }

    fs::remove_all(root);
    TEST_SUMMARY();
}
