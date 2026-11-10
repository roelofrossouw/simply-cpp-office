// Pictures out of a .docx built here: the text's order (DrawingML and the older VML), each once,
// then the ones not placed; external links skipped; raster() by name; not-a-docx errors. And a
// .docx's text: list numbers (through styles too), tables, breaks, what's left out, the options.
#include <sc_archive.h>
#include <office.h>
#include <sc_test.h>

#include <filesystem>
#include <fstream>
#include <random>
#include <sstream>
#include <string>

namespace fs = std::filesystem;

namespace {
    const std::string body = R"xml(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
</w:document>)xml";
    const std::string rels = R"xml(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.jpeg"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/word/media/image3.gif"/>
  <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.com/x.png" TargetMode="External"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>)xml";

    // An agreement as Word writes one: clause numbers from numbering.xml, set through paragraph styles
    // (Clause1/Clause2, linked to levels 0 and 1) and directly.
    const std::string agreement = R"xml(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:body>
    <w:p><w:r><w:t>DEED OF SALE</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Clause1"/></w:pPr><w:r><w:t>DEFINITIONS</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Clause2"/></w:pPr><w:r><w:t>Purchaser means</w:t></w:r><w:r><w:tab/><w:t>Ann</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>first</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>second</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Clause1"/></w:pPr><w:r><w:t>CONDITIONS PRECEDENT</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Clause2"/></w:pPr><w:r><w:t>within 30</w:t></w:r><w:del><w:r><w:delText>60</w:delText></w:r></w:del><w:r><w:t> days</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>a bullet</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>roman from five</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>legal</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Clause1"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>unnumbered</w:t></w:r></w:p>
    <w:p/><w:p/><w:p/>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>Price</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>R1 000 000</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>Erf</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>123</w:t></w:r><w:r><w:br/><w:t>Johannesburg</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p><w:r><w:t>non</w:t></w:r><w:r><w:noBreakHyphen/><w:t>refundable</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/><w:instrText> PAGE </w:instrText></w:r><w:r><w:t>7</w:t></w:r></w:p>
    <w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><w:txbxContent><w:p><w:r><w:t>in a box</w:t></w:r></w:p></w:txbxContent></w:drawing></mc:Choice><mc:Fallback><w:pict><w:txbxContent><w:p><w:r><w:t>in a box (old)</w:t></w:r></w:p></w:txbxContent></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>
    <w:sdt><w:sdtPr><w:alias w:val="Seller"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Bob Seller</w:t></w:r></w:p></w:sdtContent></w:sdt>
  </w:body>
</w:document>)xml";
    const std::string numbering = R"xml(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="10">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Clause1"/><w:lvlText w:val="%1."/></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Clause2"/><w:lvlText w:val="%1.%2"/></w:lvl>
    <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%3)"/></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="11">
    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="*"/></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="12">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1)"/></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:isLgl/><w:lvlText w:val="%1.%2"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="10"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="11"/></w:num>
  <w:num w:numId="3"><w:abstractNumId w:val="12"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>
</w:numbering>)xml";
    const std::string styles = R"xml(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="ClauseBase"><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Clause1"><w:basedOn w:val="ClauseBase"/></w:style>
  <w:style w:type="paragraph" w:styleId="Clause2"><w:basedOn w:val="ClauseBase"/></w:style>
</w:styles>)xml";
    const std::string agreement_rels = R"xml(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>)xml";
    const std::string header = R"xml(<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Initials: ____</w:t></w:r></w:p></w:hdr>)xml";
    const std::string footer = R"xml(<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Page footer</w:t></w:r></w:p></w:ftr>)xml";
    const std::string footnotes = R"xml(<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:r><w:t>See the title deed.</w:t></w:r></w:p></w:footnote></w:footnotes>)xml";

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

    const auto words = root / "agreement.docx";
    {
        auto out = sc::archive::create(words);
        out.add("[Content_Types].xml", "<Types/>");
        out.add("word/document.xml", agreement);
        out.add("word/numbering.xml", numbering);
        out.add("word/styles.xml", styles);
        out.add("word/_rels/document.xml.rels", agreement_rels);
        out.add("word/header1.xml", header);
        out.add("word/footer1.xml", footer);
        out.add("word/footnotes.xml", footnotes);
        out.close();
    }

    SECTION("Text: a line a paragraph, list numbers as Word shows them, tables as rows of cells");
    {
        const auto text = sc::office::word::text(words);
        const string expected = "DEED OF SALE\n"
                                "1. DEFINITIONS\n"                  // through its style, linked to level 0
                                "1.1 Purchaser means\tAnn\n"        // level 1 through its style
                                "(a) first\n"                        // set directly, level 2
                                "(b) second\n"
                                "2. CONDITIONS PRECEDENT\n"         // level 0 again: the levels below restart
                                "2.1 within 30 days\n"              // deleted text left out
                                "a bullet\n"                         // bullets get no number
                                "V) roman from five\n"              // startOverride, upperRoman
                                "5.1 legal\n"                        // isLgl: the levels as decimals
                                "unnumbered\n"                       // numId 0: none, though its style has one
                                "\n"                                 // three empty paragraphs: one blank line
                                "Price | R1 000 000\n"
                                "Erf | 123\nJohannesburg\n"          // a line break in a cell
                                "non-refundable7\n"                  // a non-breaking hyphen; the field code left out
                                "in a box\n\n"                       // the text box once, not its fallback copy (its own
                                                                     // paragraph inside the outer one: a blank line after)
                                "Bob Seller";                        // a content control's text
        CHECK_EQ(text, expected);
        CHECK_EQ(sc::office::word::text_in(slurp(words)), expected);
    }

    SECTION("Text: headers and footers, notes, when asked for");
    {
        const auto all = sc::office::word::text(words, {.headers_footers = true, .notes = true});
        CHECK(all.starts_with("DEED OF SALE\n1. DEFINITIONS"));
        CHECK(all.ends_with("Bob Seller\n\nInitials: ____\n\nPage footer\n\nSee the title deed."));
        CHECK(sc::office::word::text(words).find("Initials") == string::npos);
    }

    SECTION("Text: the pictures' document has its words; not a .docx throws");
    {
        CHECK_EQ(sc::office::word::text(docx), string{"Front\n\nBack"}); // the pictures' paragraphs: one blank line
        { std::ofstream{root / "old.doc"} << "\xD0\xCF\x11\xE0 an old binary .doc"; }
        CHECK_THROWS_AS((void) sc::office::word::text(root / "old.doc"), std::runtime_error);
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
