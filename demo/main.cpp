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
    doc.addParagraph("Advanced Multi-Level Lists:");
    
    // 1. Create an abstract numbering schema
    auto abstractNum = doc.numbering().addAbstractNumbering(1);
    
    openword::ListLevel lvl0;
    lvl0.levelIndex = 0;
    lvl0.format = openword::NumberingFormat::UpperRoman;
    lvl0.text = "%1.";
    abstractNum.addLevel(lvl0);
    
    openword::ListLevel lvl1;
    lvl1.levelIndex = 1;
    lvl1.format = openword::NumberingFormat::Decimal;
    lvl1.text = "%1.%2.";
    lvl1.indentTwips = 1440;
    abstractNum.addLevel(lvl1);
    
    openword::ListLevel lvl2;
    lvl2.levelIndex = 2;
    lvl2.format = openword::NumberingFormat::LowerLetter;
    lvl2.text = "(%3)";
    lvl2.indentTwips = 2160;
    abstractNum.addLevel(lvl2);

    // 2. Instantiate a concrete list from it
    int listId1 = doc.numbering().addList(1);
    
    // 3. Apply it
    doc.addParagraph("First Chapter").setList(listId1, 0);
    doc.addParagraph("Section one").setList(listId1, 1);
    doc.addParagraph("Point a").setList(listId1, 2);
    doc.addParagraph("Point b").setList(listId1, 2);
    doc.addParagraph("Section two").setList(listId1, 1);
    doc.addParagraph("Second Chapter").setList(listId1, 0);
    
    // 4. Restart numbering for a new block
    doc.addParagraph();
    doc.addParagraph("New Section (Restarted Numbering):");
    int listId2 = doc.numbering().addList(1, listId1); // Restart from listId1
    doc.addParagraph("Restarted First Chapter").setList(listId2, 0);
    doc.addParagraph("Section one").setList(listId2, 1);

    doc.save("test_07_lists.docx");
    fmt::print("- test_07_lists.docx (Advanced Multi-Level Lists)\n");
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

void test_mutability() {
    openword::Document doc;
    
    doc.addParagraph("This is the mutability and extraction demo.").setStyle("Heading1");
    
    // --- 1. Clone Row (Template looping simulation) ---
    doc.addParagraph("1. Dynamic Table Row Cloning:");
    auto t = doc.addTable(2, 3);
    t.setBorders(openword::BorderSettings{openword::BorderStyle::Single, 4, "000000"},
                 openword::BorderSettings{openword::BorderStyle::Dashed, 2, "888888"});
                 
    t.cell(0, 0).addParagraph().addRun("Item ID").setBold(true);
    t.cell(0, 1).addParagraph().addRun("Name").setBold(true);
    t.cell(0, 2).addParagraph().addRun("Price").setBold(true);
    
    // Create a template row with placeholders
    auto tmplRow = t.row(1);
    tmplRow.cells()[0].addParagraph().addRun("{{ID}}").setColor(openword::Color(255, 0, 0)); // Red ID
    tmplRow.cells()[1].addParagraph("{{NAME}}");
    tmplRow.cells()[2].addParagraph().addRun("${{PRICE}}").setBold(true);
    
    // Simulate database data
    struct Data { std::string id, name, price; };
    std::vector<Data> db = {
        {"001", "Mechanical Keyboard", "150.00"},
        {"002", "Ergonomic Mouse", "80.00"},
        {"003", "Curved Monitor", "350.00"}
    };
    
    // Loop and clone
    for (const auto& rowData : db) {
        auto newRow = tmplRow.cloneAfter();
        newRow.replaceText("{{ID}}", rowData.id);
        newRow.replaceText("{{NAME}}", rowData.name);
        newRow.replaceText("{{PRICE}}", rowData.price);
    }
    
    // Remove the placeholder template row
    tmplRow.remove();
    
    doc.addParagraph(); // spacing
    
    // --- 2. Remove Node ---
    doc.addParagraph("2. Node Removal Test:");
    auto p_keep1 = doc.addParagraph("You should see this paragraph.");
    auto p_remove = doc.addParagraph("THIS PARAGRAPH SHOULD BE INVISIBLE. IT WILL BE DELETED.");
    auto p_keep2 = doc.addParagraph("And you should see this paragraph right after the first one.");
    
    p_remove.remove(); // Nuke it from DOM
    
    doc.addParagraph(); // spacing

    // --- 3. Dynamic Insert After ---
    doc.addParagraph("3. Insertion After Middle Nodes:");
    auto first = doc.addParagraph("Step 1: Planning");
    auto third = doc.addParagraph("Step 3: Testing"); // Whoops, forgot step 2
    
    // Insert step 2 dynamically between 1 and 3
    first.insertParagraphAfter().addRun("Step 2: Implementation (Dynamically inserted!)").setColor(openword::Color(0, 0, 255)).setBold(true);
    
    doc.save("test_15_mutability.docx");
    fmt::print("- test_15_mutability.docx (AST Mutability & Looping)\n");
}

void test_advanced_layout() {
    openword::Document doc;
    
    // 1. Enable Even & Odd Headers
    doc.setEvenAndOddHeaders(true);
    
    // 2. Setup the section margins and headers
    auto sec = doc.finalSection();
    
    openword::Margins m;
    m.top = 2880;     // 2 inches
    m.bottom = 2880;  // 2 inches
    m.left = 4320;    // 3 inches (huge left margin)
    m.right = 1440;   // 1 inch
    sec.setMargins(m);
    
    // 3. Set Headers
    auto hFirst = sec.addHeader(openword::HeaderFooterType::First);
    hFirst.addParagraph().addRun("--- FIRST PAGE HEADER (Cover) ---").setBold(true);
    
    auto hEven = sec.addHeader(openword::HeaderFooterType::Even);
    hEven.addParagraph("--- EVEN PAGE HEADER ---").setAlignment("left");
    
    auto hOdd = sec.addHeader(openword::HeaderFooterType::Default);
    hOdd.addParagraph("--- ODD PAGE HEADER ---").setAlignment("right");

    // Page 1 (First)
    doc.addParagraph("This is the COVER PAGE. Notice the huge 3-inch left margin and 2-inch top margin. The header should say 'FIRST PAGE HEADER'.").setStyle("Heading1");
    // Word doesn't have an explicit 'addPageBreak' yet in our API, but a section break triggers a new page layout.
    // Or we can just insert a ton of text to spill over.
    for (int i=0; i<30; i++) doc.addParagraph("Filler text to push to next page...");

    // Page 2 (Even)
    doc.addParagraph("This is PAGE 2. The header should be on the LEFT and say 'EVEN PAGE HEADER'.").setStyle("Heading1");
    for (int i=0; i<30; i++) doc.addParagraph("Filler text to push to next page...");

    // Page 3 (Odd)
    doc.addParagraph("This is PAGE 3. The header should be on the RIGHT and say 'ODD PAGE HEADER'.").setStyle("Heading1");

    doc.save("test_16_advanced_layout.docx");
    fmt::print("- test_16_advanced_layout.docx (Margins, First/Even/Odd Headers)\n");
}

void test_resource_pool() {
    openword::Document doc;
    
    doc.addParagraph("This document contains the same image 100 times, but the generated DOCX file should only contain 1 physical image file inside the ZIP archive thanks to Resource Pooling.");
    
    auto t = doc.addTable(10, 10);
    for (int i = 0; i < 10; ++i) {
        for (int j = 0; j < 10; ++j) {
            t.cell(i, j).addParagraph().addImage("tests/test.jpg", 0.05);
        }
    }
    
    doc.save("test_17_resource_pool.docx");
    fmt::print("- test_17_resource_pool.docx (Image Hash Deduplication & Pooling)\n");
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
    test_mutability();
    test_advanced_layout();
    test_resource_pool();
    fmt::print("\nDone! Please verify test_05_tables.docx for the new advanced table layout.\n");
    return 0;
}


