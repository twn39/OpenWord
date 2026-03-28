#include <openword/Document.h>
#include <fmt/core.h>

void test_basic_text() {
    openword::Document doc;
    auto p = doc.addParagraph("This is a basic text test.");
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
    p.addRun("Big. ").setFontSize(48); // 24pt
    p.addRun("Highlight. ").setHighlight("yellow");
    p.addRun("Shade. ").setShading(openword::Color(255, 0, 0));
    p.addRun("Colored. ").setColor(openword::Color(0, 128, 255));
    p.addRun("Strike. ").setStrike(true);
    p.addRun("Subscript. ").setVertAlign(openword::VertAlign::Subscript);
    p.addRun("Arial. ").setFontFamily("Arial");

    // Test spacing and indentation
    auto p2 = doc.addParagraph("This paragraph has custom indentation and spacing. It is indented by 0.5 inches on the first line, and 0.25 inches on the left. It also has 12pt space before and 6pt space after.");
    p2.setIndentation(360, 0, 720); // 360 twips = 0.25", 720 twips = 0.5"
    p2.setSpacing(240, 120); // 240 twips = 12pt, 120 twips = 6pt

    doc.save("test_02_formatting.docx");
    fmt::print("- test_02_formatting.docx (Text Formatting, Spacing & Indentation)\n");
}

void test_styles() {
    openword::Document doc;
    doc.addStyle("MyHeading1", "Heading 1");
    auto p1 = doc.addParagraph("Styled Heading");
    p1.setStyle("MyHeading1").setAlignment("center");
    doc.addParagraph("Normal paragraph following heading.");
    doc.save("test_03_styles.docx");
    fmt::print("- test_03_styles.docx (Styles & Alignment)\n");
}

void test_image() {
    openword::Document doc;
    doc.addParagraph("Here is an image, and its size will be parsed automatically:");
    auto p1 = doc.addParagraph();
    p1.addImage("tests/test.jpg");
    
    doc.addParagraph("Here is the same image, scaled to 50%:");
    auto p2 = doc.addParagraph();
    p2.addImage("tests/test.jpg", 0.5);
    
    doc.save("test_04_image.docx");
    fmt::print("- test_04_image.docx (Images with Auto-sizing and Scaling)\n");
}

void test_tables() {
    openword::Document doc;
    doc.addParagraph("Here is a table with merged cells:");
    auto table = doc.addTable(4, 4);
    
    // Fill table
    for (int r = 0; r < 4; ++r) {
        for (int c = 0; c < 4; ++c) {
            table.cell(r, c).addParagraph(fmt::format("R{}C{}", r, c));
        }
    }
    
    // Merge first row horizontally
    table.mergeCells(0, 0, 0, 3);
    table.cell(0, 0).addParagraph("Merged Header (Span 4)").setAlignment("center");
    
    // Merge first column vertically (excluding header)
    table.mergeCells(1, 0, 3, 0);
    table.cell(1, 0).addParagraph("V-Merge (Span 3)");

    doc.save("test_05_tables.docx");
    fmt::print("- test_05_tables.docx (Tables & Cell Merging)\n");
}

int main() {
    fmt::print("Generating isolated test files...\n");
    test_basic_text();
    test_text_formatting();
    test_styles();
    test_image();
    test_tables();
    fmt::print("\nDone! Please test opening each file individually in Word.\n");
    return 0;
}
