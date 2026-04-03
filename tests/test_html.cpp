#include <catch2/catch_test_macros.hpp>
#include <openword/Document.h>
#include <iostream>

using namespace openword;

TEST_CASE("HTML to Word Rendering", "[html]") {
    Document doc;
    
    SECTION("Basic formatting") {
        doc.addHtml("<p><b>Bold</b> and <i>Italic</i> with <span style=\"color: #FF0000\">Red</span> text.</p>");
        auto els = doc.elements();
        REQUIRE(els.size() > 0);
        auto p = std::get<Paragraph>(els[0]);
        auto text = p.text();
        REQUIRE(text == "Bold and Italic with Red text.");
    }
    
    SECTION("Headings") {
        doc.addHtml("<h1>Heading 1</h1><h2>Heading 2</h2>");
        auto elements = doc.elements();
        REQUIRE(elements.size() == 3);
        REQUIRE(std::get<Paragraph>(elements[0]).text() == "Heading 1");
        REQUIRE(std::get<Paragraph>(elements[1]).text() == "Heading 2");
    }
    
    SECTION("Line breaks") {
        doc.addHtml("<p>Line 1<br>Line 2<br/>Line 3</p>");
        auto els = doc.elements();
        REQUIRE(els.size() > 0);
        auto p = std::get<Paragraph>(els[0]);
        REQUIRE(p.text() == "Line 1Line 2Line 3");
    }
    
    SECTION("Cell HTML") {
        auto t = doc.addTable(1, 1);
        t.cell(0, 0).addHtml("<b>In cell</b>");
        REQUIRE(t.cell(0, 0).text() == "In cell");
    }
}
