#include <catch2/catch_test_macros.hpp>
#include "openword/Document.h"

TEST_CASE("Structured Data Extraction", "[extraction]") {
    openword::Document doc;
    
    SECTION("Extract text from paragraphs and runs") {
        auto p = doc.addParagraph();
        p.addRun("Hello");
        p.addRun(" World!");
        
        REQUIRE(p.text() == "Hello World!");
        
        auto paragraphs = doc.paragraphs();
        REQUIRE(paragraphs.size() == 1);
        REQUIRE(paragraphs[0].text() == "Hello World!");
    }
    
    SECTION("Extract text from Tables and Cells") {
        auto t = doc.addTable(2, 2);
        t.cell(0, 0).addParagraph("Row1 Col1");
        t.cell(0, 1).addParagraph("Row1 Col2");
        t.cell(1, 0).addParagraph("Row2 Col1");
        t.cell(1, 1).addParagraph("Row2 Col2");
        
        // Test cell text
        REQUIRE(t.cell(0, 0).text() == "Row1 Col1");
        
        // Test row text
        auto row0 = t.row(0);
        auto cells = row0.cells();
        REQUIRE(cells.size() == 2);
        REQUIRE(row0.text() == "Row1 Col1 Row1 Col2");
        
        // Test table text
        auto rows = t.rows();
        REQUIRE(rows.size() == 2);
        REQUIRE(t.text() == "Row1 Col1 Row1 Col2\nRow2 Col1 Row2 Col2");
        
        // Test document tables extraction
        auto tables = doc.tables();
        REQUIRE(tables.size() == 1);
        REQUIRE(tables[0].text() == "Row1 Col1 Row1 Col2\nRow2 Col1 Row2 Col2");
    }
}
