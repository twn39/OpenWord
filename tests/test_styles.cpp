#include <catch2/catch_test_macros.hpp>
#include "openword/Document.h"
#include <filesystem>

std::string extract_file_from_zip(const std::string& zip_path, const std::string& internal_file); // from test_document

TEST_CASE("Advanced Styles Generation", "[styles]") {
    openword::Document doc;
    
    // Create a base style
    doc.styles().add("MyBaseStyle")
        .setName("My Base Style")
        .setPrimary(true)
        .setUiPriority(10);
        
    // Create a derived style
    doc.styles().add("MyDerivedStyle")
        .setName("My Derived Style")
        .setBasedOn("MyBaseStyle")
        .setNextStyle("Normal")
        .setHidden(true);
        
    auto p = doc.addParagraph("This paragraph uses the derived style.");
    p.setStyle("MyDerivedStyle");
    
    std::string filename = "test_adv_styles_extract.docx";
    REQUIRE(doc.save(filename.c_str()) == true);
    
    SECTION("styles.xml contains accurate formatting tags") {
        std::string styles_xml = extract_file_from_zip(filename, "word/styles.xml");
        
        // Base Style Verification
        REQUIRE(styles_xml.find("<w:style w:type=\"paragraph\" w:styleId=\"MyBaseStyle\">") != std::string::npos);
        REQUIRE(styles_xml.find("<w:name w:val=\"My Base Style\"/>") != std::string::npos);
        REQUIRE(styles_xml.find("<w:qFormat/>") != std::string::npos);
        REQUIRE(styles_xml.find("<w:uiPriority w:val=\"10\"/>") != std::string::npos);
        
        // Derived Style Verification
        REQUIRE(styles_xml.find("<w:style w:type=\"paragraph\" w:styleId=\"MyDerivedStyle\">") != std::string::npos);
        REQUIRE(styles_xml.find("<w:name w:val=\"My Derived Style\"/>") != std::string::npos);
        REQUIRE(styles_xml.find("<w:basedOn w:val=\"MyBaseStyle\"/>") != std::string::npos);
        REQUIRE(styles_xml.find("<w:next w:val=\"Normal\"/>") != std::string::npos);
        REQUIRE(styles_xml.find("<w:hidden/>") != std::string::npos);
        REQUIRE(styles_xml.find("<w:semiHidden/>") != std::string::npos);
    }
    
    std::filesystem::remove(filename);
}
