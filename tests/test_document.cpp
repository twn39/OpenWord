#include <catch2/catch_test_macros.hpp>
#include <openword/Document.h>
#include <pugixml.hpp>
#include <zip.h>
#include <string>

std::string extract_file_from_zip(const std::string& zip_path, const std::string& internal_file) {
    int err = 0;
    zip_t* z = zip_open(zip_path.c_str(), 0, &err);
    REQUIRE(z != nullptr);

    zip_stat_t st;
    zip_stat_init(&st);
    int res = zip_stat(z, internal_file.c_str(), 0, &st);
    REQUIRE(res == 0);

    zip_file_t* f = zip_fopen(z, internal_file.c_str(), 0);
    REQUIRE(f != nullptr);

    std::string content(st.size, '\0');
    zip_fread(f, content.data(), st.size);
    zip_fclose(f);
    zip_close(z);

    return content;
}

TEST_CASE("Document MathML to OMML conversion", "[document]") {
    openword::Document doc;
    
    SECTION("Basic MathML string is converted correctly") {
        std::string mathml = "<math><mi>x</mi></math>";
        std::string omml = doc.convertMathMLToOMML(mathml);
        
        REQUIRE_FALSE(omml.empty());
        REQUIRE(omml == "<m:oMath></m:oMath>");
    }
}

TEST_CASE("Document Structure Creation", "[document]") {
    openword::Document doc;
    
    SECTION("Can add paragraph and run with styling") {
        auto p = doc.addParagraph("Test Paragraph");
        auto r = p.addRun("Bold text");
        r.setBold(true).setFontSize(24);
        
        bool success = doc.save("test_structure.docx");
        REQUIRE(success == true);
    }
}

TEST_CASE("OOXML Structure Validation", "[ooxml]") {
    openword::Document doc;
    
    SECTION("Validates that output docx contains required OOXML files and valid XML") {
        doc.addStyle("TestStyle", "A Test Style");
        auto p = doc.addParagraph("Paragraph content");
        p.setStyle("TestStyle");
        auto r = p.addRun("Run content");
        r.setBold(true);
        
        std::string filename = "validation_test.docx";
        REQUIRE(doc.save(filename.c_str()) == true);
        
        // 1. Check [Content_Types].xml
        std::string content_types = extract_file_from_zip(filename, "[Content_Types].xml");
        pugi::xml_document ct_doc;
        REQUIRE(ct_doc.load_string(content_types.c_str()));
        REQUIRE(ct_doc.child("Types").attribute("xmlns").value() == std::string("http://schemas.openxmlformats.org/package/2006/content-types"));
        
        // 2. Check document.xml
        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        pugi::xml_document doc_dom;
        REQUIRE(doc_dom.load_string(doc_xml.c_str()));
        
        auto body = doc_dom.child("w:document").child("w:body");
        REQUIRE(body);
        auto p_node = body.child("w:p");
        REQUIRE(p_node);
        
        auto pStyle = p_node.child("w:pPr").child("w:pStyle");
        REQUIRE(pStyle.attribute("w:val").value() == std::string("TestStyle"));
        
        auto r1 = p_node.child("w:r");
        REQUIRE(r1);
        REQUIRE(r1.child("w:t").text().get() == std::string("Paragraph content"));
        
        auto r2 = r1.next_sibling("w:r");
        REQUIRE(r2);
        REQUIRE(r2.child("w:t").text().get() == std::string("Run content"));
        REQUIRE(r2.child("w:rPr").child("w:b"));

        // 3. Check styles.xml
        std::string styles_xml = extract_file_from_zip(filename, "word/styles.xml");
        pugi::xml_document styles_dom;
        REQUIRE(styles_dom.load_string(styles_xml.c_str()));
        auto style_node = styles_dom.child("w:styles").child("w:style");
        REQUIRE(style_node);
        REQUIRE(style_node.attribute("w:styleId").value() == std::string("TestStyle"));
    }
}
