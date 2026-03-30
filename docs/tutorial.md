# OpenWord Tutorial {#tutorial}

Welcome to the OpenWord tutorial. This guide demonstrates how to use the modern C++17 API to programmatically generate complex Office Open XML Word (`.docx`) documents.

## 1. Basic Text and Formatting

OpenWord uses a fluent API for text formatting. You create a paragraph, and then append `Run` objects to it. Properties can be chained continuously.

```cpp
#include <openword/Document.h>

void test_text_formatting() {
    openword::Document doc;
    auto p = doc.addParagraph();
    p.addRun("Normal. ");
    p.addRun("Bold. ").setBold(true);
    p.addRun("Italic. ").setItalic(true);
    p.addRun("Underline. ").setUnderline("single");
    p.addRun("Big. ").setFontSize(48); // 24 pt
    p.addRun("Colored. ").setColor(openword::Color(0, 128, 255));

    auto p2 = doc.addParagraph("This paragraph has custom indentation and spacing.");
    p2.setIndentation(360, 0, 720);
    p2.setSpacing(240, 120);

    doc.save("test_02_formatting.docx");
}
```

## 2. Advanced and Nested Tables

Tables are created using `Document::addTable`. You can dynamically format borders, set shading, span cells, and even create infinitely nested tables.

```cpp
void test_nested_table() {
    openword::Document doc;
    auto t1 = doc.addTable(2, 2);
    t1.cell(0, 0).addParagraph("Outer Top-Left");
    t1.cell(0, 1).addParagraph("Outer Top-Right");

    // Nested Table Creation
    auto t2 = t1.cell(1, 0).addTable(2, 2);
    t2.setBorders(openword::BorderSettings{openword::BorderStyle::Dashed, 4, "FF0000"});
    t2.cell(0, 0).addParagraph("Inner 1");
    t2.cell(0, 1).addParagraph("Inner 2");
    t2.cell(1, 0).addParagraph("Inner 3");
    t2.cell(1, 1).addParagraph("Inner 4");

    // Insert Image directly into a cell
    t1.cell(1, 1).addImage("tests/test.jpg", 0.5);

    doc.save("test_23_nested_table.docx");
}
```

## 3. Headers, Footers, and Page Numbers

You can define independent headers and footers for different document sections. Field codes (like dynamic page numbers) are fully abstracted.

```cpp
void test_headers_footers() {
    openword::Document doc;

    auto sec = doc.finalSection();

    // Add a Header
    auto header = sec.addHeader();
    header.addParagraph("Confidential Document - Company Internal");

    // Add a Footer with Dynamic Page Numbers
    auto footer = sec.addFooter();
    auto pf = footer.addParagraph();
    pf.setAlignment("center");
    pf.addRun("Page ");
    pf.addPageNumber();
    pf.addRun(" of ");
    pf.addTotalPages();

    doc.addParagraph("This is the first page (Section 1). It has a header and a footer.");
    for (int i = 0; i < 15; ++i) {
        doc.addParagraph("Filler text to pad out page 1...");
    }

    // New Section (Remove header, keep footer)
    auto sec2 = doc.addParagraph().appendSectionBreak();
    sec2.removeHeader(); // Unlink header
    doc.addParagraph("This is Section 2 (Starting on Page 2). The header should be GONE here, but the footer should continue.");

    doc.save("test_25_headers_footers.docx");
}
```

## 4. Line Breaks and Page Breaks

Use `addLineBreak()` for a soft return (Shift+Enter) within the same paragraph, and `addPageBreak()` to force pagination.

```cpp
void test_breaks() {
    openword::Document doc;

    // Line Breaks
    auto p1 = doc.addParagraph();
    p1.addRun("First Line");
    p1.addRun().addLineBreak(); // Soft return
    p1.addRun("Second Line within the same paragraph");

    // Page Break
    auto p2 = doc.addParagraph("This is the last sentence on page 1.");
    p2.addRun().addPageBreak();

    doc.addParagraph("This sentence starts exactly on page 2.");

    doc.save("test_24_breaks.docx");
}
```

## 5. Comments

You can programmatically attach comments (with author information) to specific runs of text.

```cpp
void test_comments() {
    openword::Document doc;
    
    // Register comment first
    int cid = doc.createComment("This is a test comment!", "Dr. Developer", "DD");
    
    auto p = doc.addParagraph("This is some regular text, and ");
    
    // Chain the anchor ID to a specific run
    p.addRun("here is the text with a comment attached").addComment(cid);
    p.addRun(". And here is more normal text.");

    doc.save("test_22_comments.docx");
}
```

## 6. Styles Inheritance and Overrides

You can create named styles that inherit from built-in styles, and even manipulate global document defaults.

```cpp
void test_styles() {
    openword::Document doc;

    // 1. Built-in styles
    auto p1 = doc.addParagraph("Standard Built-in Heading 1");
    p1.setStyle("Heading1");

    // 2. Custom Style with Inheritance and UI Controls
    auto customTheme = doc.styles().add("MyCustomTheme");
    customTheme.setName("My Custom Theme")
        .setBasedOn("Heading1") // Inherits fonts and sizes from Heading1
        .setNextStyle("Normal") // Hitting enter goes back to normal text
        .setPrimary(true)       // Show in Quick Styles gallery
        .setUiPriority(1);      // Force it to the front

    customTheme.getFont().setColor("0000FF");

    auto p2 = doc.addParagraph("This is a custom style inheriting from Heading1 but dyed Blue.");
    p2.setStyle("MyCustomTheme");

    // 3. Document Defaults Override (docDefaults)
    doc.styles().getDefaultFont().setSize(28).setName("Arial").setColor("333333");
    
    doc.addParagraph("This text has NO style applied directly to it. It uses the global document defaults (14pt Arial).");

    doc.save("test_03_styles.docx");
}
```

## 7. Metadata and Custom Properties

Attach core and extended properties, and define strongly-typed custom parameters visible in the Word properties panel.

```cpp
void test_metadata_advanced() {
    openword::Document doc;

    openword::Metadata meta;
    meta.title = "Generated Report";
    meta.author = "System Engine";

    // Strongly-typed custom properties
    meta.customProperties["ApprovalStatus"] = openword::CustomProperty{openword::CustomProperty::Type::Text, "Approved"};
    meta.customProperties["ConfidenceScore"] = openword::CustomProperty{openword::CustomProperty::Type::Double, "99.9"};

    doc.setMetadata(meta);
    doc.save("test_21_advanced_metadata.docx");
}
```
