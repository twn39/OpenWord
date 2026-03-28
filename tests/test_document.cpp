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

TEST_CASE("Document Loading and Parsing", "[load]") {
    std::string filename = "parse_test.docx";
    
    // Phase 1: Create a document
    {
        openword::Document doc;
        auto p1 = doc.addParagraph("First paragraph text.");
        auto p2 = doc.addParagraph("Second ");
        auto r = p2.addRun("paragraph text.");
        r.setBold(true);
        REQUIRE(doc.save(filename.c_str()));
    }

    // Phase 2: Load the document and verify contents recursively
    {
        openword::Document doc2;
        REQUIRE(doc2.load(filename.c_str()));
        
        auto paras = doc2.paragraphs();
        REQUIRE(paras.size() == 2);
        
        // Verify text aggregation
        REQUIRE(paras[0].text() == "First paragraph text.");
        REQUIRE(paras[1].text() == "Second paragraph text.");
        
        // Verify structural runs count
        REQUIRE(paras[1].runs().size() == 2);
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

TEST_CASE("Text Formatting Validation", "[formatting]") {
    openword::Document doc;
    
    SECTION("Can set various text formatting properties") {
        auto p = doc.addParagraph();
        p.addRun("Bold").setBold(true);
        p.addRun("Italic").setItalic(true);
        p.addRun("Underline").setUnderline("single");
        p.addRun("Size").setFontSize(48);
        p.addRun("Highlight").setHighlight("yellow");
        p.addRun("Shading").setShading(openword::Color(255, 0, 0));
        p.addRun("Colored").setColor(openword::Color(0, 255, 0));
        p.addRun("Font").setFontFamily("Arial");
        p.addRun("Strike").setStrike(true);
        p.addRun("Subscript").setVertAlign(openword::VertAlign::Subscript);

        std::string filename = "test_formatting.docx";
        REQUIRE(doc.save(filename.c_str()) == true);

        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        pugi::xml_document doc_dom;
        REQUIRE(doc_dom.load_string(doc_xml.c_str()));
        
        auto body = doc_dom.child("w:document").child("w:body");
        auto p_node = body.child("w:p");
        REQUIRE(p_node);
        
        auto r = p_node.child("w:r");
        REQUIRE(r.child("w:t").text().get() == std::string("Bold"));
        REQUIRE(r.child("w:rPr").child("w:b"));

        r = r.next_sibling("w:r");
        REQUIRE(r.child("w:t").text().get() == std::string("Italic"));
        REQUIRE(r.child("w:rPr").child("w:i"));

        r = r.next_sibling("w:r");
        REQUIRE(r.child("w:t").text().get() == std::string("Underline"));
        REQUIRE(r.child("w:rPr").child("w:u").attribute("w:val").value() == std::string("single"));

        r = r.next_sibling("w:r");
        REQUIRE(r.child("w:t").text().get() == std::string("Size"));
        REQUIRE(r.child("w:rPr").child("w:sz").attribute("w:val").value() == std::string("48"));

        r = r.next_sibling("w:r");
        REQUIRE(r.child("w:t").text().get() == std::string("Highlight"));
        REQUIRE(r.child("w:rPr").child("w:highlight").attribute("w:val").value() == std::string("yellow"));

        r = r.next_sibling("w:r");
        REQUIRE(r.child("w:t").text().get() == std::string("Shading"));
        REQUIRE(r.child("w:rPr").child("w:shd").attribute("w:fill").value() == std::string("FF0000"));

        r = r.next_sibling("w:r");
        REQUIRE(r.child("w:t").text().get() == std::string("Colored"));
        REQUIRE(r.child("w:rPr").child("w:color").attribute("w:val").value() == std::string("00FF00"));

        r = r.next_sibling("w:r");
        REQUIRE(r.child("w:t").text().get() == std::string("Font"));
        REQUIRE(r.child("w:rPr").child("w:rFonts").attribute("w:ascii").value() == std::string("Arial"));

        r = r.next_sibling("w:r");
        REQUIRE(r.child("w:t").text().get() == std::string("Strike"));
        REQUIRE(r.child("w:rPr").child("w:strike"));

        r = r.next_sibling("w:r");
        REQUIRE(r.child("w:t").text().get() == std::string("Subscript"));
        REQUIRE(r.child("w:rPr").child("w:vertAlign").attribute("w:val").value() == std::string("subscript"));
    }
}

TEST_CASE("Table Creation and Validation", "[table]") {
    openword::Document doc;
    
    SECTION("Can create a 2x2 table and write into cells") {
        auto table = doc.addTable(2, 2);
        
        auto cell00 = table.cell(0, 0);
        cell00.addParagraph("Top Left");

        auto cell11 = table.cell(1, 1);
        auto p11 = cell11.addParagraph();
        p11.addRun("Bottom Right").setBold(true);

        std::string filename = "test_table.docx";
        REQUIRE(doc.save(filename.c_str()) == true);

        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        pugi::xml_document doc_dom;
        REQUIRE(doc_dom.load_string(doc_xml.c_str()));
        
        auto body = doc_dom.child("w:document").child("w:body");
        auto tbl = body.child("w:tbl");
        REQUIRE(tbl);

        int row_count = 0;
        for (auto tr : tbl.children("w:tr")) {
            row_count++;
            int col_count = 0;
            for (auto tc : tr.children("w:tc")) {
                col_count++;
                if (row_count == 1 && col_count == 1) {
                    REQUIRE(tc.child("w:p").child("w:r").child("w:t").text().get() == std::string("Top Left"));
                }
                if (row_count == 2 && col_count == 2) {
                    REQUIRE(tc.child("w:p").child("w:r").child("w:t").text().get() == std::string("Bottom Right"));
                    REQUIRE(tc.child("w:p").child("w:r").child("w:rPr").child("w:b"));
                }
            }
            REQUIRE(col_count == 2);
        }
        REQUIRE(row_count == 2);
        
        auto last_node = body.last_child();
        REQUIRE(std::string(last_node.name()) == "w:sectPr");
    }
}

TEST_CASE("Image Embedding and Validation", "[image]") {
    openword::Document doc;
    
    SECTION("Image can be embedded, creating valid DrawingML and relationships") {
        std::string dummy_img = "dummy_test_image.jpg";
        FILE* f = fopen(dummy_img.c_str(), "wb");
        if (f) {
            fputs("fake image data", f);
            fclose(f);
        }

        auto p = doc.addParagraph("Image below:");
        p.addImage(dummy_img.c_str());

        std::string filename = "test_image.docx";
        REQUIRE(doc.save(filename.c_str()) == true);

        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        pugi::xml_document doc_dom;
        REQUIRE(doc_dom.load_string(doc_xml.c_str()));
        
        auto body = doc_dom.child("w:document").child("w:body");
        
        pugi::xml_node drawing_node;
        for (auto para : body.children("w:p")) {
            for (auto run : para.children("w:r")) {
                if (run.child("w:drawing")) {
                    drawing_node = run.child("w:drawing");
                    break;
                }
            }
        }
        REQUIRE(drawing_node);
        
        auto inline_node = drawing_node.child("wp:inline");
        REQUIRE(inline_node);
        
        auto extent = inline_node.child("wp:extent");
        REQUIRE(extent);
        REQUIRE(std::string(extent.attribute("cx").value()) == "4572000");

        auto docPr = inline_node.child("wp:docPr");
        REQUIRE(docPr);
        REQUIRE(std::string(docPr.attribute("name").value()) == "Picture 1");

        auto graphic = inline_node.child("a:graphic");
        auto pic = graphic.child("a:graphicData").child("pic:pic");
        REQUIRE(pic);
        
        auto blip = pic.child("pic:blipFill").child("a:blip");
        REQUIRE(blip);
        std::string rId = blip.attribute("r:embed").value();
        REQUIRE_FALSE(rId.empty());
        
        std::string rels_xml = extract_file_from_zip(filename, "word/_rels/document.xml.rels");
        pugi::xml_document rels_dom;
        REQUIRE(rels_dom.load_string(rels_xml.c_str()));
        
        bool found_rel = false;
        std::string target_path = "";
        for (auto rel : rels_dom.child("Relationships").children("Relationship")) {
            if (std::string(rel.attribute("Id").value()) == rId) {
                found_rel = true;
                target_path = rel.attribute("Target").value();
                REQUIRE(std::string(rel.attribute("Type").value()) == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
            }
        }
        REQUIRE(found_rel == true);
        REQUIRE(target_path == "media/image1.jpg");

        std::string media_file = extract_file_from_zip(filename, "word/media/image1.jpg");
        REQUIRE(media_file.size() == std::string("fake image data").size()); 

        remove(dummy_img.c_str());
    }
}

TEST_CASE("Paragraph Spacing and Indentation Validation", "[formatting]") {
    openword::Document doc;
    
    SECTION("Can set spacing and indentation") {
        auto p = doc.addParagraph("Indented and spaced text");
        p.setIndentation(360, 0, 720); // 0.25 left indent, 0.5 first line indent
        p.setSpacing(240, 120, 360, "exact"); // 12pt before, 6pt after, exactly 18pt line spacing

        std::string filename = "test_paragraph_format.docx";
        REQUIRE(doc.save(filename.c_str()) == true);

        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        pugi::xml_document doc_dom;
        REQUIRE(doc_dom.load_string(doc_xml.c_str()));
        
        auto body = doc_dom.child("w:document").child("w:body");
        auto p_node = body.child("w:p");
        auto pPr = p_node.child("w:pPr");
        
        auto ind = pPr.child("w:ind");
        REQUIRE(ind);
        REQUIRE(ind.attribute("w:left").value() == std::string("360"));
        REQUIRE(ind.attribute("w:firstLine").value() == std::string("720"));
        
        auto spacing = pPr.child("w:spacing");
        REQUIRE(spacing);
        REQUIRE(spacing.attribute("w:before").value() == std::string("240"));
        REQUIRE(spacing.attribute("w:after").value() == std::string("120"));
        REQUIRE(spacing.attribute("w:line").value() == std::string("360"));
        REQUIRE(spacing.attribute("w:lineRule").value() == std::string("exact"));
    }
}

TEST_CASE("Table Merging Validation", "[table]") {
    openword::Document doc;
    
    SECTION("Can horizontally and vertically merge cells") {
        auto table = doc.addTable(3, 3);
        
        // Merge top row horizontally
        table.mergeCells(0, 0, 0, 2);
        
        // Merge first column vertically in the remaining rows
        table.mergeCells(1, 0, 2, 0);

        std::string filename = "test_table_merge.docx";
        REQUIRE(doc.save(filename.c_str()) == true);

        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        pugi::xml_document doc_dom;
        REQUIRE(doc_dom.load_string(doc_xml.c_str()));
        
        auto body = doc_dom.child("w:document").child("w:body");
        auto tbl = body.child("w:tbl");
        
        // Check row 0
        auto tr0 = tbl.child("w:tr");
        auto tc00 = tr0.child("w:tc");
        REQUIRE(tc00.child("w:tcPr").child("w:gridSpan").attribute("w:val").value() == std::string("3"));
        REQUIRE(tc00.next_sibling("w:tc") == nullptr); // Ensure other cells deleted
        
        // Check row 1
        auto tr1 = tr0.next_sibling("w:tr");
        auto tc10 = tr1.child("w:tc");
        REQUIRE(tc10.child("w:tcPr").child("w:vMerge").attribute("w:val").value() == std::string("restart"));
        auto tc11 = tc10.next_sibling("w:tc");
        REQUIRE(tc11); // Col 1
        auto tc12 = tc11.next_sibling("w:tc");
        REQUIRE(tc12); // Col 2
        
        // Check row 2
        auto tr2 = tr1.next_sibling("w:tr");
        auto tc20 = tr2.child("w:tc");
        REQUIRE(tc20.child("w:tcPr").child("w:vMerge").attribute("w:val").value() == std::string("continue"));
    }
}

TEST_CASE("Image Scaling Validation", "[image]") {
    openword::Document doc;
    
    SECTION("Image sizes match provided scale factor") {
        std::string dummy_img = "dummy_scale_image.png";
        FILE* f = fopen(dummy_img.c_str(), "wb");
        if (f) {
            // Write a fake 100x50 PNG signature to be correctly parsed
            unsigned char fake_png[24] = {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, 
                0x00, 0x00, 0x00, 0x64, // Width: 100
                0x00, 0x00, 0x00, 0x32  // Height: 50
            };
            fwrite(fake_png, 1, 24, f);
            fclose(f);
        }

        // Test normal scale (1.0)
        auto p1 = doc.addParagraph();
        p1.addImage(dummy_img.c_str());
        
        // Test 0.5 scale
        auto p2 = doc.addParagraph();
        p2.addImage(dummy_img.c_str(), 0.5);

        std::string filename = "test_image_scale.docx";
        REQUIRE(doc.save(filename.c_str()) == true);

        std::string doc_xml = extract_file_from_zip(filename, "word/document.xml");
        pugi::xml_document doc_dom;
        REQUIRE(doc_dom.load_string(doc_xml.c_str()));
        
        auto body = doc_dom.child("w:document").child("w:body");
        auto p_iter = body.children("w:p").begin();
        
        // Check normal image size (100 * 9525 = 952500, 50 * 9525 = 476250)
        auto img1 = p_iter->child("w:r").child("w:drawing").child("wp:inline").child("wp:extent");
        REQUIRE(std::string(img1.attribute("cx").value()) == "952500");
        REQUIRE(std::string(img1.attribute("cy").value()) == "476250");

        ++p_iter;
        
        // Check 0.5 scaled image size
        auto img2 = p_iter->child("w:r").child("w:drawing").child("wp:inline").child("wp:extent");
        REQUIRE(std::string(img2.attribute("cx").value()) == "476250");
        REQUIRE(std::string(img2.attribute("cy").value()) == "238125");

        remove(dummy_img.c_str());
    }
}
