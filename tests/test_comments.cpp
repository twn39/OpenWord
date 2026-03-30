#include <catch2/catch_test_macros.hpp>
#include "openword/Document.h"
#include <filesystem>

std::string extract_file_from_zip(const std::string& zip_path, const std::string& internal_file); // from test_document

TEST_CASE("Comment Generation", "[comments]") {
    openword::Document doc;
    
    int c1 = doc.createComment("First Comment", "Author1", "A1");
    int c2 = doc.createComment("Second Comment", "Author2", "A2");
    
    auto p = doc.addParagraph("Hello ");
    p.addRun("World").addComment(c1);
    p.addRun(" Test").addComment(c2);
    
    std::string filename = "test_comment_extract.docx";
    REQUIRE(doc.save(filename.c_str()) == true);
    
    SECTION("comments.xml is generated") {
        std::string comments_xml = extract_file_from_zip(filename, "word/comments.xml");
        REQUIRE(comments_xml.find("<w:comment w:id=\"0\" w:author=\"Author1\" w:initials=\"A1\">") != std::string::npos);
        REQUIRE(comments_xml.find("<w:t>First Comment</w:t>") != std::string::npos);
        REQUIRE(comments_xml.find("<w:comment w:id=\"1\" w:author=\"Author2\" w:initials=\"A2\">") != std::string::npos);
        REQUIRE(comments_xml.find("<w:t>Second Comment</w:t>") != std::string::npos);
    }
    
    SECTION("document.xml contains anchors") {
        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        REQUIRE(doc_xml.find("<w:commentRangeStart w:id=\"0\"/>") != std::string::npos);
        REQUIRE(doc_xml.find("<w:t>World</w:t>") != std::string::npos);
        REQUIRE(doc_xml.find("<w:commentRangeEnd w:id=\"0\"/>") != std::string::npos);
        REQUIRE(doc_xml.find("<w:commentReference w:id=\"0\"/>") != std::string::npos);
        
        REQUIRE(doc_xml.find("<w:commentRangeStart w:id=\"1\"/>") != std::string::npos);
        REQUIRE(doc_xml.find("<w:t xml:space=\"preserve\"> Test</w:t>") != std::string::npos);
        REQUIRE(doc_xml.find("<w:commentRangeEnd w:id=\"1\"/>") != std::string::npos);
        REQUIRE(doc_xml.find("<w:commentReference w:id=\"1\"/>") != std::string::npos);
    }
    
    SECTION("rels are correctly mapped") {
        std::string rels_xml = extract_file_from_zip(filename, "word/_rels/document.xml.rels");
        REQUIRE(rels_xml.find("Target=\"comments.xml\"") != std::string::npos);
    }
    
    std::filesystem::remove(filename);
}
