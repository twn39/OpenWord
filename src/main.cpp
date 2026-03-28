#include <openword/Document.h>
#include <fmt/core.h>

void generate_minimal() {
    openword::Document doc;
    doc.addParagraph("01 - Minimal Document (No styles, no rels)");
    doc.addParagraph("This file tests the most basic XML structure.");
    doc.save("01_minimal.docx");
}

void generate_styles_defined() {
    openword::Document doc;
    doc.addStyle("MyStyle", "Custom Style Name");
    doc.addParagraph("02 - Styles Defined (Not applied)");
    doc.addParagraph("This file includes styles.xml but doesn't use the style yet.");
    doc.save("02_styles_defined.docx");
}

void generate_styles_applied() {
    openword::Document doc;
    doc.addStyle("MyHeading1", "Heading 1");
    auto p = doc.addParagraph("03 - Styles Applied");
    p.setStyle("MyHeading1");
    doc.addParagraph("This file includes styles.xml and applies MyHeading1 to the paragraph above.");
    doc.save("03_styles_applied.docx");
}

int main() {
    fmt::print("Generating diagnostic files...\n");

    generate_minimal();
    fmt::print("- Generated: 01_minimal.docx\n");

    generate_styles_defined();
    fmt::print("- Generated: 02_styles_defined.docx\n");

    generate_styles_applied();
    fmt::print("- Generated: 03_styles_applied.docx\n");

    fmt::print("Done. Please check which files Word can open.\n");
    return 0;
}
