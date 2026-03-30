#include <catch2/catch_test_macros.hpp>
#include "openword/Document.h"
#include <filesystem>
#include <zip.h>
#include <string>

std::string extract_file_from_zip(const std::string& zip_path, const std::string& internal_file); // Defined in test_document.cpp


TEST_CASE("Metadata and Custom Properties parsing", "[metadata]") {
    openword::Document doc;
    
    openword::Metadata meta;
    meta.title = "AST Parser Design";
    meta.author = "John Doe";
    meta.subject = "Architecture";
    meta.keywords = "C++, OOXML, AST";
    meta.comments = "This is a test description.";
    meta.lastModifiedBy = "Jane Doe";
    meta.category = "Design Document";
    meta.company = "Tech Corp";
    meta.manager = "Boss";
    meta.hyperlinkBase = "https://example.com";
    
    meta.customProperties["DocVersion"] = openword::CustomProperty{openword::CustomProperty::Type::Text, "v1.0.0"};
    meta.customProperties["IsDraft"] = openword::CustomProperty{openword::CustomProperty::Type::Boolean, "true"};
    meta.customProperties["RevisionCount"] = openword::CustomProperty{openword::CustomProperty::Type::Integer, "42"};
    
    doc.setMetadata(meta);
    
    std::string filename = "test_metadata_full.docx";
    REQUIRE(doc.save(filename.c_str()) == true);
    
    SECTION("Core Properties") {
        std::string core_xml = extract_file_from_zip(filename, "docProps/core.xml");
        REQUIRE(core_xml.find("<dc:title>AST Parser Design</dc:title>") != std::string::npos);
        REQUIRE(core_xml.find("<dc:creator>John Doe</dc:creator>") != std::string::npos);
        REQUIRE(core_xml.find("<dc:subject>Architecture</dc:subject>") != std::string::npos);
        REQUIRE(core_xml.find("<cp:keywords>C++, OOXML, AST</cp:keywords>") != std::string::npos);
        REQUIRE(core_xml.find("<dc:description>This is a test description.</dc:description>") != std::string::npos);
        REQUIRE(core_xml.find("<cp:lastModifiedBy>Jane Doe</cp:lastModifiedBy>") != std::string::npos);
        REQUIRE(core_xml.find("<cp:category>Design Document</cp:category>") != std::string::npos);
    }
    
    SECTION("Extended App Properties") {
        std::string app_xml = extract_file_from_zip(filename, "docProps/app.xml");
        REQUIRE(app_xml.find("<Company>Tech Corp</Company>") != std::string::npos);
        REQUIRE(app_xml.find("<Manager>Boss</Manager>") != std::string::npos);
        REQUIRE(app_xml.find("<HyperlinkBase>https://example.com</HyperlinkBase>") != std::string::npos);
        REQUIRE(app_xml.find("<HeadingPairs>") != std::string::npos);
        REQUIRE(app_xml.find("<TitlesOfParts>") != std::string::npos);
        REQUIRE(app_xml.find("<vt:lpstr>AST Parser Design</vt:lpstr>") != std::string::npos);
    }
    
    SECTION("Custom Properties") {
        std::string custom_xml = extract_file_from_zip(filename, "docProps/custom.xml");
        REQUIRE(custom_xml.find("name=\"DocVersion\"") != std::string::npos);
        REQUIRE(custom_xml.find("<vt:lpwstr>v1.0.0</vt:lpwstr>") != std::string::npos);
        
        REQUIRE(custom_xml.find("name=\"IsDraft\"") != std::string::npos);
        REQUIRE(custom_xml.find("<vt:bool>true</vt:bool>") != std::string::npos);
        
        REQUIRE(custom_xml.find("name=\"RevisionCount\"") != std::string::npos);
        REQUIRE(custom_xml.find("<vt:i4>42</vt:i4>") != std::string::npos);
    }
    
    SECTION("Rels Integration") {
        std::string rels_xml = extract_file_from_zip(filename, "_rels/.rels");
        REQUIRE(rels_xml.find("Target=\"docProps/core.xml\"") != std::string::npos);
        REQUIRE(rels_xml.find("Target=\"docProps/app.xml\"") != std::string::npos);
        REQUIRE(rels_xml.find("Target=\"docProps/custom.xml\"") != std::string::npos);
    }
    
    std::filesystem::remove(filename);
}
