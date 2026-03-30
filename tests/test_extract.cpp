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
    
    
    SECTION("Polymorphic AST Traversal and Styling Extractors") {
        doc.addParagraph("Heading Block").setStyle("Heading1").setAlignment("center");
        doc.addParagraph("- List Item 1").setList(1, 0);
        auto t = doc.addTable(1, 1);
        t.cell(0,0).addParagraph("Cell Content");
        
        auto p_run = doc.addParagraph();
        p_run.addRun("Bold text").setBold(true);
        p_run.addRun("Italic").setItalic(true);
        
        auto elements = doc.elements();
        // size might be 5 if there's a trailing w:sectPr
        REQUIRE(elements.size() >= 4);
        
        // 1. First element: Paragraph (Heading1, Center)
        REQUIRE(std::holds_alternative<openword::Paragraph>(elements[0]));
        auto p1 = std::get<openword::Paragraph>(elements[0]);
        REQUIRE(p1.styleId() == "Heading1");
        REQUIRE(p1.alignment() == "center");
        REQUIRE(p1.isList() == false);
        
        // 2. Second element: Paragraph (List)
        REQUIRE(std::holds_alternative<openword::Paragraph>(elements[1]));
        auto p2 = std::get<openword::Paragraph>(elements[1]);
        REQUIRE(p2.isList() == true);
        REQUIRE(p2.listLevel() == 0);
        
        // 3. Third element: Table
        REQUIRE(std::holds_alternative<openword::Table>(elements[2]));
        auto tableNode = std::get<openword::Table>(elements[2]);
        auto cellNode = tableNode.cell(0, 0);
        REQUIRE(cellNode.gridSpan() == 1);
        REQUIRE(cellNode.elements().size() == 1); // contains 1 paragraph
        
        // 4. Fourth element: Paragraph (Runs analysis)
        auto p4 = std::get<openword::Paragraph>(elements[3]);
        auto runs = p4.runs();
        REQUIRE(runs.size() == 2);
        REQUIRE(runs[0].isBold() == true);
        REQUIRE(runs[0].isItalic() == false);
        REQUIRE(runs[0].text() == "Bold text");
        
        REQUIRE(runs[1].isBold() == false);
        REQUIRE(runs[1].isItalic() == true);
        REQUIRE(runs[1].text() == "Italic");
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
