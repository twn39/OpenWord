#include <openword/Document.h>
#include <fmt/core.h>
#include <vector>
#include <string>

void test_basic_text() {
    openword::Document doc;
    doc.addParagraph("This is a basic text test.");
    doc.save("test_01_basic.docx");
    fmt::print("- test_01_basic.docx (Basic Paragraph)\n");
}

void test_text_formatting() {
    openword::Document doc;
    auto p = doc.addParagraph();
    p.addRun("Normal. ");
    p.addRun("Bold. ").setBold(true);
    p.addRun("Italic. ").setItalic(true);
    p.addRun("Underline. ").setUnderline("single");
    p.addRun("Big. ").setFontSize(48); 
    p.addRun("Colored. ").setColor(openword::Color(0, 128, 255));

    auto p2 = doc.addParagraph("This paragraph has custom indentation and spacing.");
    p2.setIndentation(360, 0, 720);
    p2.setSpacing(240, 120);

    doc.save("test_02_formatting.docx");
    fmt::print("- test_02_formatting.docx (Text Formatting)\n");
}

void test_styles() {
    openword::Document doc;
    auto p1 = doc.addParagraph("Standard Built-in Heading 1");
    p1.setStyle("Heading1");
    auto p2 = doc.addParagraph("Standard Built-in Heading 2");
    p2.setStyle("Heading2");
    doc.save("test_03_styles.docx");
    fmt::print("- test_03_styles.docx (Styles)\n");
}

void test_image() {
    openword::Document doc;
    doc.addParagraph("Images test:");
    doc.addParagraph().addImage("tests/test.jpg", 0.5);
    doc.save("test_04_image.docx");
    fmt::print("- test_04_image.docx (Images)\n");
}

void test_tables() {
    openword::Document doc;
    doc.addParagraph("Advanced Table test:");
    auto table = doc.addTable(3, 3);
    
    // 1. Column Widths (Ergonomic API)
    table.setColumnWidths({2000, 3000, 4000});
    
    // 2. Row Heights
    table.row(0).setHeight(800, openword::HeightRule::Exact);
    
    // 3. Borders (Ergonomic Outer/Inner setup)
    openword::BorderSettings outer{openword::BorderStyle::Thick, 12, "000000"};
    openword::BorderSettings inner{openword::BorderStyle::Dashed, 4, "888888"};
    table.setBorders(outer, inner);
    
    // 4. Merge Cells & Alignment
    table.mergeCells(0, 0, 0, 2);
    auto headerCell = table.cell(0, 0);
    headerCell.setShading("E7E6E6"); // Light gray background
    headerCell.setVertAlign(openword::VerticalAlignment::Center);
    headerCell.addParagraph("Merged Header (Gray, Centered)").setAlignment("center");
    
    // 5. Fill some data
    table.cell(1, 0).addParagraph("Row 2, Col 1").setAlignment("center");
    table.cell(1, 1).addParagraph("Row 2, Col 2\nMultiple lines\nVertical Center");
    table.cell(1, 1).setVertAlign(openword::VerticalAlignment::Center);
    table.cell(1, 2).addParagraph("Bottom aligned");
    table.cell(1, 2).setVertAlign(openword::VerticalAlignment::Bottom);
    
    doc.save("test_05_tables.docx");
    fmt::print("- test_05_tables.docx (Tables)\n");
}

void test_sections_and_headers() {
    openword::Document doc;
    doc.finalSection().setPageSize(16838, 11906, openword::Orientation::Landscape);
    doc.addParagraph("Landscape content");
    doc.save("test_06_sections.docx");
    fmt::print("- test_06_sections.docx (Sections)\n");
}

void test_lists() {
    openword::Document doc;
    doc.addParagraph("List test:");
    doc.addParagraph("Item 1").setList(openword::ListType::Bullet, 0);
    doc.save("test_07_lists.docx");
    fmt::print("- test_07_lists.docx (Lists)\n");
}

void test_math() {
    openword::Document doc;
    doc.addParagraph("LaTeX Comprehensive Capability Test").setStyle("Heading1");

    auto add_latex = [&](const std::string& label, const std::string& latex) {
        doc.addParagraph(label + " (" + latex + "):");
        std::string omml = doc.convertLaTeXToOMML(latex);
        if (!omml.empty()) {
            doc.addParagraph().addEquation(omml);
        } else {
            fmt::print("- Critical failure (empty output) for: {}\n", latex);
        }
    };

    // Math tests...
    add_latex("Greek Variants", "\\varepsilon \\varkappa \\varpi \\varrho \\varsigma \\varphi");
    add_latex("Integrals", "\\int_a^b x^2 dx \\quad \\iint_D f(x,y) dS");
    add_latex("Matrices (Standard)", "\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}");

    doc.save("test_08_math.docx");
    fmt::print("- test_08_math.docx (LaTeX Full Capability Showcase Generated)\n");
}

void test_toc_and_metadata() {
    openword::Document doc;
    
    openword::Metadata meta;
    meta.title = "OpenWord Capabilities Demo";
    meta.author = "Antigravity";
    meta.subject = "C++ DOCX Generation";
    doc.setMetadata(meta);

    doc.addTableOfContents("Document Table of Contents", 3);
    
    auto p1 = doc.addParagraph();
    p1.setStyle("Heading1");
    p1.addRun("Chapter 1: Document Structure");
    doc.addParagraph("This document demonstrates TOC generation and metadata injection.");
    
    doc.save("test_09_toc_metadata.docx");
    fmt::print("- test_09_toc_metadata.docx (TOC and Metadata)\n");
}

void test_links_and_references() {
    openword::Document doc;
    
    auto p = doc.addParagraph("This paragraph contains multiple links: ");
    p.addHyperlink("Google", "https://google.com");
    p.addRun(" and ");
    p.addHyperlink("GitHub", "https://github.com");
    p.addRun(". ");
    
    p.insertBookmark("MyBookmark");
    p.addRun("You can jump back here using ");
    p.addInternalLink("this internal link", "MyBookmark");
    p.addRun(".");
    
    int fn1 = doc.createFootnote("This is a footnote reference.");
    int en1 = doc.createEndnote("This is an endnote reference at the end of the document.");
    
    auto p2 = doc.addParagraph("This sentence has a footnote");
    p2.addFootnoteReference(fn1);
    p2.addRun(" and an endnote");
    p2.addEndnoteReference(en1);
    p2.addRun(".");
    
    doc.save("test_10_links_refs.docx");
    fmt::print("- test_10_links_refs.docx (Links, Bookmarks, Footnotes, Endnotes)\n");
}

void test_columns_and_whitespace() {
    openword::Document doc;
    
    auto p1 = doc.addParagraph();
    p1.addRun("This   text   preserves   consecutive   spaces   because   of   xml:space=\"preserve\".");
    
    auto p2 = doc.addParagraph("This paragraph starts a two-column section.");
    auto sec = p2.appendSectionBreak();
    sec.setColumns(2, 720); // 2 columns, 0.5 inch spacing (720 twips)
    
    for (int i = 0; i < 10; ++i) {
        doc.addParagraph(fmt::format("Column text line {}. This text should wrap within the narrow column constraints.", i + 1));
    }
    
    doc.save("test_11_columns_space.docx");
    fmt::print("- test_11_columns_space.docx (Columns and Whitespace)\n");
}

void test_replace() {
    openword::Document doc;
    
    auto p1 = doc.addParagraph();
    p1.addRun("Dear ");
    p1.addRun("{{");
    p1.addRun("NAME").setBold(true).setColor(openword::Color(0, 0, 255)); // Blue
    p1.addRun("}}");
    p1.addRun(", thank you for your purchase!");

    auto t = doc.addTable(2, 2);
    t.cell(0, 0).addParagraph("Invoice Date:");
    t.cell(0, 1).addParagraph("{{DATE}}");
    t.cell(1, 0).addParagraph("Amount Due:");
    t.cell(1, 1).addParagraph("${{AMOUNT}}");

    // Replace text globally across paragraph runs and table cells
    doc.replaceText("{{NAME}}", "Jane Doe");
    doc.replaceText("{{DATE}}", "2024-10-01");
    doc.replaceText("{{AMOUNT}}", "9,999.00");

    doc.save("test_12_replace_text.docx");
    fmt::print("- test_12_replace_text.docx (Text Replacement Engine)\n");
}

void test_floating_image() {
    openword::Document doc;
    
    // 1. Standard Inline Image
    doc.addParagraph("1. This is a standard inline image (behaves like a large text character):");
    doc.addParagraph().addImage("tests/test.jpg", 0.3);
    
    doc.addParagraph();
    
    // 2. Floating Image (Behind Text)
    auto p = doc.addParagraph();
    
    // X Offset: 1 inch = 914400 EMU. Let's put it 2 inches from column start.
    // Y Offset: 0.5 inches = 457200 EMU from paragraph start.
    p.addImage("tests/test.jpg", 0.35, openword::ImagePosition::BehindText, 3200400, 4114800);
    
    for (int i = 0; i < 5; ++i) {
        p.addRun("2. This text is flowing over the floating image! Because the image is set to float BEHIND the text. This is extremely useful for watermarks or stamping corporate seals on contracts. ");
    }
    
    doc.save("test_14_floating_image.docx");
    fmt::print("- test_14_floating_image.docx (Floating Images)\n");
}

int main() {
    fmt::print("Generating capability test files...\n");
    test_basic_text();
    test_text_formatting();
    test_styles();
    test_image();
    test_tables();
    test_sections_and_headers();
    test_lists();
    test_math();
    test_toc_and_metadata();
    test_links_and_references();
    test_columns_and_whitespace();
    test_replace();
    test_floating_image();
    fmt::print("\nDone! Please verify test_05_tables.docx for the new advanced table layout.\n");
    return 0;
}


