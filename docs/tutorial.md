# OpenWord Tutorial {#tutorial}

Welcome to the OpenWord tutorial. This guide demonstrates how to use the modern C++17 API to programmatically generate complex Office Open XML Word (`.docx`) documents.

## 1. Basic Text and Formatting

OpenWord uses a fluent API for text formatting. You create a paragraph, and then append `Run` objects to it. Properties can be chained continuously.

```cpp
#include <openword/Document.h>

int main() {
    openword::Document doc;
    
    auto p = doc.addParagraph();
    p.addRun("Normal text. ");
    p.addRun("Bold and Red! ").setBold(true).setColor(openword::Color(255, 0, 0));
    p.addRun("Big text.").setFontSize(48); // 24 pt (sizes are in half-points)
    
    doc.save("formatting.docx");
    return 0;
}
```

## 2. Advanced Tables

Tables are created using `Document::addTable`. You can dynamically format borders, set shading, span cells, and even create infinitely nested tables.

```cpp
auto t = doc.addTable(2, 2);

// Apply ergonomic borders
t.setBorders(openword::BorderSettings{openword::BorderStyle::Dashed, 4, "0000FF"});

// Insert text and alignment
t.cell(0, 0).addParagraph("Cell 1");
t.cell(0, 1).addParagraph("Cell 2").setAlignment("center");

// Nested Table Creation
auto innerTable = t.cell(1, 0).addTable(2, 2);
innerTable.cell(0, 0).addParagraph("Inner Cell");
```

## 3. Headers, Footers, and Page Numbers

You can define independent headers and footers for different document sections. Field codes (like dynamic page numbers) are fully abstracted.

```cpp
auto sec = doc.finalSection();

// Add a Header
auto header = sec.addHeader();
header.addParagraph("Confidential Internal Document");

// Add a Footer with Dynamic Page Numbers
auto footer = sec.addFooter();
auto pf = footer.addParagraph();
pf.setAlignment("center");
pf.addRun("Page ");
pf.addPageNumber();
pf.addRun(" of ");
pf.addTotalPages();
```

## 4. Lists and Numbering

Creating fully compliant OOXML bulleted or numbered lists is effortless through the `NumberingCollection`.

```cpp
int bulletId = doc.numbering().addBulletList();

doc.addParagraph("First Item").setList(bulletId, 0); // Level 0
doc.addParagraph("Sub-item A").setList(bulletId, 1); // Level 1 (indented)
doc.addParagraph("Sub-item B").setList(bulletId, 1);
doc.addParagraph("Second Item").setList(bulletId, 0);
```

## 5. Metadata and Custom Properties

Attach core and extended properties, and define strongly-typed custom parameters visible in the Word properties panel.

```cpp
openword::Metadata meta;
meta.title = "Generated Report";
meta.author = "System Engine";

// Strongly-typed custom properties
meta.customProperties["ApprovalStatus"] = openword::CustomProperty{openword::CustomProperty::Type::Text, "Approved"};
meta.customProperties["ConfidenceScore"] = openword::CustomProperty{openword::CustomProperty::Type::Double, "99.9"};

doc.setMetadata(meta);
```
