#include <catch2/catch_test_macros.hpp>
#include <openword/Document.h>
#include <pugixml.hpp>
#include <zip.h>
#include <string>
#include <filesystem>
#include <fstream>
#include <sstream>
#include <algorithm>

/**
 * @brief Helper to extract a file from the generated .docx (ZIP)
 */
std::string extract_file_from_zip(const std::string& zip_path, const std::string& internal_file) {
    int err = 0;
    zip_t* z = zip_open(zip_path.c_str(), 0, &err);
    if (!z) return "";

    zip_stat_t st;
    zip_stat_init(&st);
    int res = zip_stat(z, internal_file.c_str(), 0, &st);
    if (res != 0) {
        zip_close(z);
        return "";
    }

    zip_file_t* f = zip_fopen(z, internal_file.c_str(), 0);
    if (!f) {
        zip_close(z);
        return "";
    }

    std::string content(st.size, '\0');
    zip_fread(f, content.data(), st.size);
    zip_fclose(f);
    zip_close(z);

    return content;
}

TEST_CASE("Lists and Numbering Validation", "[lists]") {
    openword::Document doc;
    
    SECTION("Create and validate bullet and numbered lists") {
        doc.addParagraph("Bullet Item 1").setList(openword::ListType::Bullet, 0);
        doc.addParagraph("Numbered Item 1").setList(openword::ListType::Numbered, 0);
        doc.addParagraph("Sub-bullet").setList(openword::ListType::Bullet, 1);

        std::string filename = "test_lists_val.docx";
        REQUIRE(doc.save(filename.c_str()));

        // 1. Verify numbering.xml existence and structure
        std::string num_xml = extract_file_from_zip(filename, "word/numbering.xml");
        REQUIRE_FALSE(num_xml.empty());
        
        pugi::xml_document num_doc;
        num_doc.load_string(num_xml.c_str());
        auto root = num_doc.child("w:numbering");
        REQUIRE(root);
        
        // Should have at least 2 abstractNum and 2 num nodes
        REQUIRE(std::distance(root.children("w:abstractNum").begin(), root.children("w:abstractNum").end()) >= 2);
        REQUIRE(std::distance(root.children("w:num").begin(), root.children("w:num").end()) >= 2);

        // Verify w:lvl nesting (CRITICAL for Word compatibility)
        auto absNum0 = root.child("w:abstractNum");
        REQUIRE(absNum0.child("w:lvl"));
        REQUIRE(absNum0.child("w:lvl").attribute("w:ilvl").as_int() == 0);

        // 2. Verify document.xml references
        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        pugi::xml_document xml_doc;
        xml_doc.load_string(doc_xml.c_str());
        
        auto body = xml_doc.child("w:document").child("w:body");
        auto p1 = body.child("w:p");
        auto numPr1 = p1.child("w:pPr").child("w:numPr");
        REQUIRE(numPr1);
        REQUIRE(numPr1.child("w:numId").attribute("w:val").as_int() == 1); // Bullet
        
        auto p2 = p1.next_sibling("w:p");
        auto numPr2 = p2.child("w:pPr").child("w:numPr");
        REQUIRE(numPr2);
        REQUIRE(numPr2.child("w:numId").attribute("w:val").as_int() == 2); // Numbered

        // 3. Verify Relationships
        std::string rels_xml = extract_file_from_zip(filename, "word/_rels/document.xml.rels");
        REQUIRE(rels_xml.find("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\"") != std::string::npos);
        REQUIRE(rels_xml.find("Target=\"numbering.xml\"") != std::string::npos);

        // 4. Verify Content_Types
        std::string ct_xml = extract_file_from_zip(filename, "[Content_Types].xml");
        REQUIRE(ct_xml.find("PartName=\"/word/numbering.xml\"") != std::string::npos);

        std::filesystem::remove(filename);
    }
}

TEST_CASE("Document Math Conversion and Insertion", "[math]") {
    openword::Document doc;
    
    SECTION("LaTeX to OMML conversion and insertion") {
        std::string latex = "\\frac{a}{b}";
        std::string omml = doc.convertLaTeXToOMML(latex);
        REQUIRE_FALSE(omml.empty());
        REQUIRE(omml.find("m:oMath") != std::string::npos);

        auto p = doc.addParagraph("Formula below:");
        auto p_math = doc.addParagraph();
        p_math.addEquation(omml);

        std::string filename = "test_math_insertion.docx";
        REQUIRE(doc.save(filename.c_str()));

        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        pugi::xml_document xml_doc;
        xml_doc.load_string(doc_xml.c_str());

        // Verify that m:oMath is inside a paragraph
        auto body = xml_doc.child("w:document").child("w:body");
        bool found_math = false;
        for (auto p_node : body.children("w:p")) {
            if (p_node.child("m:oMath")) {
                found_math = true;
                // Verify structure of the fraction
                REQUIRE(p_node.child("m:oMath").child("m:f")); 
                break;
            }
        }
        REQUIRE(found_math);
        std::filesystem::remove(filename);
    }
}

TEST_CASE("Sections and Headers/Footers Validation", "[sections]") {
    openword::Document doc;
    
    SECTION("Complex sections with headers and footers") {
        // Section 1: Landscape
        auto s1 = doc.finalSection();
        s1.setPageSize(16838, 11906, openword::Orientation::Landscape);
        auto h1 = s1.addHeader();
        h1.addParagraph("Header Section 1");
        doc.addParagraph("Content Page 1");

        // Break to Section 2: Portrait
        auto p1 = doc.addParagraph("Last para of section 1");
        auto s2 = p1.appendSectionBreak();
        s2.setPageSize(11906, 16838, openword::Orientation::Portrait);
        auto f2 = s2.addFooter();
        f2.addParagraph("Footer Section 2");
        doc.addParagraph("Content Page 2");

        std::string filename = "test_complex_sections.docx";
        REQUIRE(doc.save(filename.c_str()));

        // 1. Verify ZIP Structure
        REQUIRE_FALSE(extract_file_from_zip(filename, "word/header100.xml").empty());
        REQUIRE_FALSE(extract_file_from_zip(filename, "word/footer101.xml").empty());

        // 2. Verify Relationships
        std::string rels_xml = extract_file_from_zip(filename, "word/_rels/document.xml.rels");
        pugi::xml_document rels_doc;
        rels_doc.load_string(rels_xml.c_str());
        auto rels = rels_doc.child("Relationships");
        
        auto find_rel = [&](const std::string& target) {
            for (auto r : rels.children("Relationship")) {
                if (std::string(r.attribute("Target").value()) == target) return true;
            }
            return false;
        };
        REQUIRE(find_rel("header100.xml"));
        REQUIRE(find_rel("footer101.xml"));

        // 3. Verify Content_Types
        std::string ct_xml = extract_file_from_zip(filename, "[Content_Types].xml");
        REQUIRE(ct_xml.find("PartName=\"/word/header100.xml\"") != std::string::npos);
        REQUIRE(ct_xml.find("ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"") != std::string::npos);

        // 4. Verify XML Schema Order in document.xml
        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        pugi::xml_document xml_doc;
        xml_doc.load_string(doc_xml.c_str());
        
        // Check first section break (inside paragraph)
        auto body = xml_doc.child("w:document").child("w:body");
        auto p1_node = body.find_child_by_attribute("w:p", "w:rsidR", ""); // Just find first p with sectPr
        // Better way: find the p that has a sectPr
        pugi::xml_node sectPr1;
        for(auto p : body.children("w:p")) {
            if (p.child("w:pPr").child("w:sectPr")) {
                sectPr1 = p.child("w:pPr").child("w:sectPr");
                break;
            }
        }
        REQUIRE(sectPr1);
        
        // CRITICAL: headerReference or footerReference MUST be before pgSz
        auto first_child = sectPr1.first_child();
        std::string first_name = first_child.name();
        REQUIRE((first_name == "w:headerReference" || first_name == "w:footerReference"));
        
        auto pgSz1 = sectPr1.child("w:pgSz");
        REQUIRE(pgSz1);
        REQUIRE(std::string(pgSz1.attribute("w:orient").value()) == "landscape");

        // Check final section
        auto finalSect = body.child("w:sectPr");
        REQUIRE(finalSect);
        REQUIRE(finalSect.child("w:footerReference"));
        REQUIRE(std::string(finalSect.first_child().name()) == "w:footerReference");

        std::filesystem::remove(filename);
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
        std::filesystem::remove("test_structure.docx");
    }
}

TEST_CASE("Document Loading and Parsing", "[load]") {
    std::string filename = "parse_test.docx";
    
    {
        openword::Document doc;
        auto p1 = doc.addParagraph("First paragraph text.");
        auto p2 = doc.addParagraph("Second ");
        auto r = p2.addRun("paragraph text.");
        r.setBold(true);
        REQUIRE(doc.save(filename.c_str()));
    }

    {
        openword::Document doc2;
        REQUIRE(doc2.load(filename.c_str()));
        
        auto paras = doc2.paragraphs();
        REQUIRE(paras.size() == 2);
        REQUIRE(paras[0].text() == "First paragraph text.");
        REQUIRE(paras[1].text() == "Second paragraph text.");
        REQUIRE(paras[1].runs().size() == 2);
    }
    std::filesystem::remove(filename);
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
        
        std::string content_types = extract_file_from_zip(filename, "[Content_Types].xml");
        pugi::xml_document ct_doc;
        REQUIRE(ct_doc.load_string(content_types.c_str()));
        REQUIRE(ct_doc.child("Types").attribute("xmlns").value() == std::string("http://schemas.openxmlformats.org/package/2006/content-types"));
        
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
        REQUIRE(r1.child("w:t").text().get() == std::string("Paragraph content"));
        
        auto r2 = r1.next_sibling("w:r");
        REQUIRE(r2.child("w:t").text().get() == std::string("Run content"));
        REQUIRE(r2.child("w:rPr").child("w:b"));

        std::filesystem::remove(filename);
    }
}

TEST_CASE("Text Formatting Validation", "[formatting]") {
    openword::Document doc;
    
    SECTION("Can set various text formatting properties") {
        auto p = doc.addParagraph();
        p.addRun("Bold").setBold(true);
        p.addRun("Size").setFontSize(48);
        p.addRun("Colored").setColor(openword::Color(0, 255, 0));

        std::string filename = "test_formatting_val.docx";
        REQUIRE(doc.save(filename.c_str()) == true);

        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        pugi::xml_document doc_dom;
        REQUIRE(doc_dom.load_string(doc_xml.c_str()));
        
        auto body = doc_dom.child("w:document").child("w:body");
        auto p_node = body.child("w:p");
        
        auto r = p_node.child("w:r");
        REQUIRE(r.child("w:rPr").child("w:b"));

        r = r.next_sibling("w:r");
        REQUIRE(r.child("w:rPr").child("w:sz").attribute("w:val").value() == std::string("48"));

        std::filesystem::remove(filename);
    }
}

TEST_CASE("Table Creation and Validation", "[table]") {
    openword::Document doc;
    auto table = doc.addTable(2, 2);
    table.cell(0, 0).addParagraph("Top Left");

    std::string filename = "test_table_val.docx";
    REQUIRE(doc.save(filename.c_str()) == true);

    std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
    pugi::xml_document doc_dom;
    doc_dom.load_string(doc_xml.c_str());
    
    auto tbl = doc_dom.child("w:document").child("w:body").child("w:tbl");
    REQUIRE(tbl);
    REQUIRE(tbl.child("w:tr").child("w:tc").child("w:p").child("w:r").child("w:t").text().get() == std::string("Top Left"));

    std::filesystem::remove(filename);
}

TEST_CASE("UTF-8 Content Validation", "[utf8]") {
    openword::Document doc;
    std::string text_content = "UTF-8 中文测试 🌟";
    doc.addParagraph(text_content);

    std::string filename = "test_utf8_val.docx";
    REQUIRE(doc.save(filename.c_str()) == true);

    std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
    pugi::xml_document doc_dom;
    doc_dom.load_string(doc_xml.c_str());
    
    auto t = doc_dom.child("w:document").child("w:body").child("w:p").child("w:r").child("w:t");
    REQUIRE(t.text().get() == text_content);

    std::filesystem::remove(filename);
}

TEST_CASE("Advanced Document Features Validation", "[advanced]") {
    openword::Document doc;
    
    SECTION("Metadata and Core Properties") {
        openword::Metadata meta;
        meta.title = "Test Document API";
        meta.author = "Antigravity";
        meta.subject = "C++ Library";
        doc.setMetadata(meta);
        
        std::string filename = "test_adv_meta.docx";
        REQUIRE(doc.save(filename.c_str()) == true);

        std::string core_xml = extract_file_from_zip(filename, "docProps/core.xml");
        REQUIRE_FALSE(core_xml.empty());
        pugi::xml_document xml;
        xml.load_string(core_xml.c_str());
        auto cp = xml.child("cp:coreProperties");
        REQUIRE(std::string(cp.child("dc:title").text().get()) == "Test Document API");
        REQUIRE(std::string(cp.child("dc:creator").text().get()) == "Antigravity");

        std::filesystem::remove(filename);
    }
    
    SECTION("Hyperlinks and Multiple ID Management") {
        auto p = doc.addParagraph("Check this:");
        p.addHyperlink(" Google", "https://google.com");
        p.addHyperlink(" GitHub", "https://github.com");
        p.insertBookmark("MyBookmark");
        p.addInternalLink(" LinkBack", "MyBookmark");
        
        std::string filename = "test_adv_links.docx";
        REQUIRE(doc.save(filename.c_str()) == true);

        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        REQUIRE(doc_xml.find("w:hyperlink") != std::string::npos);
        REQUIRE(doc_xml.find("Google") != std::string::npos);
        REQUIRE(doc_xml.find("GitHub") != std::string::npos);
        REQUIRE(doc_xml.find("w:bookmarkStart") != std::string::npos);
        
        std::string rel_xml = extract_file_from_zip(filename, "word/_rels/document.xml.rels");
        REQUIRE(rel_xml.find("https://google.com") != std::string::npos);
        REQUIRE(rel_xml.find("https://github.com") != std::string::npos);

        std::filesystem::remove(filename);
    }

    SECTION("Table of Contents, Outline Levels, and Settings") {
        doc.addTableOfContents("My TOC", 3);
        
        auto p1 = doc.addParagraph();
        p1.addRun("Chapter 1");
        p1.setOutlineLevel(0);
        
        auto p2 = doc.addParagraph();
        p2.addRun("Chapter 1.1");
        p2.setOutlineLevel(1);
        
        std::string filename = "test_adv_toc.docx";
        REQUIRE(doc.save(filename.c_str()) == true);

        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        REQUIRE(doc_xml.find("w:sdt") != std::string::npos);
        REQUIRE(doc_xml.find("w:instrText") != std::string::npos);
        REQUIRE(doc_xml.find("TOC \\o") != std::string::npos);
        REQUIRE(doc_xml.find("w:outlineLvl w:val=\"0\"") != std::string::npos);
        REQUIRE(doc_xml.find("w:outlineLvl w:val=\"1\"") != std::string::npos);

        std::string settings_xml = extract_file_from_zip(filename, "word/settings.xml");
        REQUIRE(settings_xml.find("w:updateFields w:val=\"true\"") != std::string::npos);

        std::filesystem::remove(filename);
    }

    SECTION("Whitespace Preservation") {
        auto p = doc.addParagraph();
        p.addRun("NoSpace");
        p.addRun(" LeadingSpace");
        p.addRun("TrailingSpace ");

        std::string filename = "test_adv_space.docx";
        REQUIRE(doc.save(filename.c_str()) == true);
        
        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        REQUIRE(doc_xml.find("xml:space=\"preserve\"") != std::string::npos);
        REQUIRE(doc_xml.find("> LeadingSpace<") != std::string::npos);
        REQUIRE(doc_xml.find(">TrailingSpace <") != std::string::npos);
        
        std::filesystem::remove(filename);
    }

    SECTION("Footnotes and Endnotes") {
        int fnId = doc.createFootnote("This is a footnote text.");
        auto p = doc.addParagraph("Text with note");
        p.addFootnoteReference(fnId);
        
        std::string filename = "test_adv_notes.docx";
        REQUIRE(doc.save(filename.c_str()) == true);

        std::string fn_xml = extract_file_from_zip(filename, "word/footnotes.xml");
        REQUIRE_FALSE(fn_xml.empty());
        REQUIRE(fn_xml.find("This is a footnote text.") != std::string::npos);
        
        std::string rel_xml = extract_file_from_zip(filename, "word/_rels/document.xml.rels");
        REQUIRE(rel_xml.find("footnotes.xml") != std::string::npos);
        
        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        REQUIRE(doc_xml.find("w:footnoteReference") != std::string::npos);

        std::filesystem::remove(filename);
    }
    
    SECTION("Section Columns") {
        auto p = doc.addParagraph("Col 1...");
        auto s = p.appendSectionBreak();
        s.setColumns(2, 700);
        
        std::string filename = "test_adv_cols.docx";
        REQUIRE(doc.save(filename.c_str()) == true);
        
        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        REQUIRE(doc_xml.find("w:cols") != std::string::npos);
        REQUIRE(doc_xml.find("w:num=\"2\"") != std::string::npos);

        std::filesystem::remove(filename);
    }
}
