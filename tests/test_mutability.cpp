#include <catch2/catch_test_macros.hpp>
#include "openword/Document.h"

TEST_CASE("DOM Mutability and Cloning", "[mutability]") {
    openword::Document doc;
    
    SECTION("Remove a Paragraph") {
        auto p1 = doc.addParagraph("First");
        auto p2 = doc.addParagraph("Second");
        auto p3 = doc.addParagraph("Third");
        
        // Remove the middle paragraph
        p2.remove();
        
        auto paras = doc.paragraphs();
        REQUIRE(paras.size() == 2);
        REQUIRE(paras[0].text() == "First");
        REQUIRE(paras[1].text() == "Third");
    }
    
    SECTION("Clone a Paragraph After") {
        auto p1 = doc.addParagraph("Template Paragraph");
        p1.setIndentation(720, 0, 0, 0); // Give it some style to ensure style is cloned
        
        auto p2 = p1.cloneAfter();
        REQUIRE(p2.text() == "Template Paragraph");
        
        auto paras = doc.paragraphs();
        REQUIRE(paras.size() == 2);
    }
    
    SECTION("Remove and Clone Tables and Rows") {
        auto t1 = doc.addTable(2, 2);
        t1.cell(0, 0).addParagraph("A");
        t1.cell(0, 1).addParagraph("B");
        t1.cell(1, 0).addParagraph("C");
        t1.cell(1, 1).addParagraph("D");
        
        auto t2 = doc.addTable(1, 1);
        t2.cell(0, 0).addParagraph("Remove Me");
        
        REQUIRE(doc.tables().size() == 2);
        
        // Remove entire table
        t2.remove();
        REQUIRE(doc.tables().size() == 1);
        
        // Clone row
        auto newRow = t1.row(1).cloneAfter();
        REQUIRE(t1.rows().size() == 3);
        REQUIRE(newRow.cells()[0].text() == "C");
        
        // Remove row
        t1.row(0).remove();
        REQUIRE(t1.rows().size() == 2);
        REQUIRE(t1.row(0).text() == "C D"); // Row 1 shifted up to Row 0
    }
    
    SECTION("Insert Table after Paragraph") {
        auto p1 = doc.addParagraph("Before Table");
        auto p2 = doc.addParagraph("After Table");
        
        auto t = p1.insertTableAfter(2, 2);
        t.cell(0, 0).addParagraph("Inserted Table Cell");
        
        auto paras = doc.paragraphs();
        REQUIRE(paras[0].text() == "Before Table");
        // Due to tree ordering, the table is between the paragraphs
        REQUIRE(paras[1].text() == "After Table"); 
        
        auto tables = doc.tables();
        REQUIRE(tables.size() == 1);
        REQUIRE(tables[0].cell(0, 0).text() == "Inserted Table Cell");
    }
}
