#include "openword/Document.h"
#include "openword/MarkdownConverter.h"
#include <catch2/catch_test_macros.hpp>

TEST_CASE("Markdown Conversion", "[markdown]") {
    openword::Document doc;

    SECTION("Convert simple paragraphs and headings") {
        doc.addParagraph("Heading 1").setStyle("Heading1");
        doc.addParagraph("This is a simple paragraph.");

        auto p_run = doc.addParagraph();
        p_run.addRun("Bold text ").setBold(true);
        p_run.addRun("and ");
        p_run.addRun("Italic").setItalic(true);

        doc.addParagraph("List item 1").setList(1, 0);
        doc.addParagraph("List item 2").setList(1, 0);

        std::string md = openword::convertToMarkdown(doc);

        REQUIRE(md.find("# Heading 1") != std::string::npos);
        REQUIRE(md.find("This is a simple paragraph.") != std::string::npos);
        REQUIRE(md.find("**Bold text **and *Italic*") != std::string::npos);
        REQUIRE(md.find("- List item 1") != std::string::npos);
        REQUIRE(md.find("- List item 2") != std::string::npos);
    }

    SECTION("Convert tables") {
        auto t = doc.addTable(2, 2);
        t.cell(0, 0).addParagraph("Header 1");
        t.cell(0, 1).addParagraph("Header 2");
        t.cell(1, 0).addParagraph("Row 1 Col 1");
        t.cell(1, 1).addParagraph("Row 1 Col 2");

        std::string md = openword::convertToMarkdown(doc);

        REQUIRE(md.find("| Header 1 | Header 2 |") != std::string::npos);
        REQUIRE(md.find("|---|---|") != std::string::npos);
        REQUIRE(md.find("| Row 1 Col 1 | Row 1 Col 2 |") != std::string::npos);
    }
}

TEST_CASE("Markdown Footnotes and Endnotes", "[markdown]") {
    openword::Document doc;
    int fnId = doc.createFootnote("This is a footnote text.");
    int enId = doc.createEndnote("This is an endnote text.");

    auto p = doc.addParagraph("This statement has a footnote");
    p.addFootnoteReference(fnId);
    p.addRun(" and an endnote");
    p.addEndnoteReference(enId);
    p.addRun(".");

    std::string md = openword::convertToMarkdown(doc);

    REQUIRE(md.find("This statement has a footnote[^1] and an endnote[^e_1].") != std::string::npos);
    REQUIRE(md.find("[^1]: This is a footnote text.") != std::string::npos);
    REQUIRE(md.find("[^e_1]: This is an endnote text.") != std::string::npos);
}
