#include <openword/Document.h>
#include <openword/Validator.h>
#include <iostream>
#include <vector>
#include <map>
#include <fstream>

using namespace openword;

int main() {
    Document doc;
    
    Paragraph p1 = doc.addParagraph("OpenWord Advanced Features Demo");
    p1.addRun().setBold(true);
    
    doc.addParagraph("This document was automatically generated to demonstrate the new Template Engine (cloneRowAndSetValues) and the SchemaValidator interfaces.");
    doc.addParagraph("");
    
    Paragraph p2 = doc.addParagraph("1. Template Engine Demonstration");
    p2.addRun().setBold(true);
    
    doc.addParagraph("Below is a table where the second row was originally a template containing placeholders (${name}, ${dept}, ${role}). The cloneRowAndSetValues function was called to generate 4 data rows, automatically removing the template row.");
    doc.addParagraph("");

    // Create a table with 2 rows (1 header, 1 template)
    auto table = doc.addTable(2, 3);
    
    // Setup header
    table.row(0).setCantSplit(true).setRepeatHeaderRow(true);
    table.cell(0, 0).paragraphs()[0].addRun("Employee Name").setBold(true);
    table.cell(0, 1).paragraphs()[0].addRun("Department").setBold(true);
    table.cell(0, 2).paragraphs()[0].addRun("Role").setBold(true);
    
    // Setup template row
    table.cell(1, 0).paragraphs()[0].addRun("${name}");
    table.cell(1, 1).paragraphs()[0].addRun("${dept}");
    table.cell(1, 2).paragraphs()[0].addRun("${role}");
    
    // Prepare data
    std::vector<std::map<std::string, std::string>> employees = {
        {{"${name}", "Alice Smith"}, {"${dept}", "Engineering"}, {"${role}", "Senior Developer"}},
        {{"${name}", "Bob Jones"}, {"${dept}", "Marketing"}, {"${role}", "SEO Specialist"}},
        {{"${name}", "Charlie Brown"}, {"${dept}", "Sales"}, {"${role}", "Account Executive"}},
        {{"${name}", "Diana Prince"}, {"${dept}", "Management"}, {"${role}", "CEO"}}
    };
    
    // Apply template
    int generatedRows = doc.cloneRowAndSetValues("${name}", employees);
    std::cout << "Generated " << generatedRows << " rows from the template.\n";
    
    doc.addParagraph("");
    Paragraph p3 = doc.addParagraph("2. Schema Validator Demonstration");
    p3.addRun().setBold(true);
    doc.addParagraph("The application also tested validating the generated document.xml against a dynamic XSD schema using libxml2 in the background. Note: Full OOXML validation requires the complete ECMA-376 schema files.");

    // Schema Validation (Dummy schema just to show it triggers the libxml2 engine correctly)
    const char* xsd_path = "dummy.xsd";
    std::ofstream out(xsd_path);
    out << "<?xml version=\"1.0\"?><xs:schema xmlns:xs=\"http://www.w3.org/2001/XMLSchema\"><xs:element name=\"document\"><xs:complexType><xs:sequence><xs:any processContents=\"skip\" minOccurs=\"0\" maxOccurs=\"unbounded\"/></xs:sequence><xs:anyAttribute processContents=\"skip\"/></xs:complexType></xs:element></xs:schema>";
    out.close();

    SchemaValidator validator(xsd_path);
    std::string errs;
    bool isValid = doc.validate("document.xml", validator, errs);
    std::cout << "Document validation against dummy schema executed. Result: " << (isValid ? "Passed" : "Failed (expected due to complex OOXML namespaces mismatching the dummy schema)") << "\n";
    if (!isValid) {
        std::cout << "Sample validation errors caught by libxml2:\n" << errs << "\n";
    }
    
    std::remove(xsd_path);

    // Save document
    std::string outputPath = "Template_Engine_Demo.docx";
    if (doc.save(outputPath.c_str())) {
        std::cout << "\n✅ Successfully saved document to: " << outputPath << "\n";
    } else {
        std::cerr << "\n❌ Failed to save document.\n";
    }
    
    return 0;
}
