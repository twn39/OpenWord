#include <catch2/catch_test_macros.hpp>
#include <cstdio>
#include <fstream>
#include <openword/Document.h>
#include <openword/Validator.h>

using namespace openword;

TEST_CASE("SchemaValidator basic operations", "[validator]") {
    const char *xsd_path = "test_schema.xsd";

    {
        std::ofstream out(xsd_path);
        out << "<?xml version=\"1.0\"?>"
            << "<xs:schema xmlns:xs=\"http://www.w3.org/2001/XMLSchema\">"
            << "  <xs:element name=\"note\">"
            << "    <xs:complexType>"
            << "      <xs:sequence>"
            << "        <xs:element name=\"to\" type=\"xs:string\"/>"
            << "        <xs:element name=\"from\" type=\"xs:string\"/>"
            << "      </xs:sequence>"
            << "    </xs:complexType>"
            << "  </xs:element>"
            << "</xs:schema>";
    }

    SECTION("Valid XML") {
        SchemaValidator validator(xsd_path);
        REQUIRE(validator.isValid());

        std::string xml = "<note><to>Tove</to><from>Jani</from></note>";
        std::string errs;
        bool result = validator.validate(xml.c_str(), errs);
        REQUIRE(result == true);
        REQUIRE(errs.empty());
    }

    SECTION("Invalid XML") {
        SchemaValidator validator(xsd_path);
        REQUIRE(validator.isValid());

        std::string xml = "<note><to>Tove</to><from>Jani</from><extra>Invalid</extra></note>";
        std::string errs;
        bool result = validator.validate(xml.c_str(), errs);
        REQUIRE(result == false);
        REQUIRE(!errs.empty());
    }

    SECTION("Invalid XSD path") {
        SchemaValidator validator("non_existent.xsd");
        REQUIRE(validator.isValid() == false);
    }

    std::remove(xsd_path);
}

TEST_CASE("Document::validate operations", "[validator][document]") {
    const char *xsd_path = "dummy_doc_schema.xsd";
    {
        std::ofstream out(xsd_path);
        out << "<?xml version=\"1.0\"?>"
            << "<xs:schema xmlns:xs=\"http://www.w3.org/2001/XMLSchema\">"
            << "  <xs:element name=\"document\">"
            << "    <xs:complexType>"
            << "      <xs:sequence>"
            << "        <xs:any processContents=\"skip\" minOccurs=\"0\" maxOccurs=\"unbounded\"/>"
            << "      </xs:sequence>"
            << "      <xs:anyAttribute processContents=\"skip\"/>"
            << "    </xs:complexType>"
            << "  </xs:element>"
            << "</xs:schema>";
    }

    SchemaValidator validator(xsd_path);
    REQUIRE(validator.isValid());

    Document doc;
    std::string errs;

    bool result = doc.validate("document.xml", validator, errs);

    REQUIRE(result == false);
    REQUIRE(!errs.empty());

    bool res2 = doc.validate("unknown_part.xml", validator, errs);
    REQUIRE(res2 == false);
    REQUIRE(errs.find("Unknown or unsupported") != std::string::npos);

    std::remove(xsd_path);
}
