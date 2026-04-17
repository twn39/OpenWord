#include <catch2/catch_test_macros.hpp>
#include <openword/Document.h>
#include <openword/Validator.h>
#include <zip.h>
#include <fstream>

using namespace openword;

TEST_CASE("Chart Creation and XML Structure", "[chart]") {
    Document doc;
    ChartSeries s1{"Data 1", {"A", "B"}, {}, {1.0, 2.0}, "FF0000"};

    ChartOptions options;
    options.title = "Test Chart Title";
    options.showDataLabels = true;
    options.legendPos = LegendPosition::Left;
    options.widthTwips = 6000;
    options.heightTwips = 4000;

    auto chart = doc.addChart(ChartType::Pie, {s1}, options);
    
    // Add another chart to verify indexing
    ChartSeries s2{"Data 2", {"X"}, {}, {5.0}, "00FF00"};
    doc.addChart(ChartType::Bar, {s2});

    auto charts = doc.charts();
    REQUIRE(charts.size() == 2);
    REQUIRE(charts[0].relId() != "");
    REQUIRE(charts[1].relId() != "");
    REQUIRE(charts[0].relId() != charts[1].relId());

    REQUIRE(doc.save("test_chart_output.docx") == true);

    int error = 0;
    zip_t *z = zip_open("test_chart_output.docx", ZIP_RDONLY, &error);
    REQUIRE(z != nullptr);

    zip_stat_t st;
    zip_stat_init(&st);
    REQUIRE(zip_stat(z, "word/charts/chart1.xml", 0, &st) == 0);
    REQUIRE(st.size > 0);

    zip_file_t *f_chart = zip_fopen(z, "word/charts/chart1.xml", 0);
    std::string chart_content(st.size, '\0');
    zip_fread(f_chart, chart_content.data(), st.size);
    zip_fclose(f_chart);

    // Verify injected style data and structure for Pie Chart
    REQUIRE(chart_content.find("Test Chart Title") != std::string::npos);
    REQUIRE(chart_content.find("FF0000") != std::string::npos);
    REQUIRE(chart_content.find("<c:showVal val=\"1\"") != std::string::npos); // showDataLabels
    REQUIRE(chart_content.find("<c:legendPos val=\"l\"") != std::string::npos);
    REQUIRE(chart_content.find("<c:pieChart>") != std::string::npos);
    
    // Verify XML Namespaces are present on the chartSpace node
    REQUIRE(chart_content.find("xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"") != std::string::npos);
    REQUIRE(chart_content.find("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"") != std::string::npos);

    REQUIRE(zip_stat(z, "word/document.xml", 0, &st) == 0);
    zip_file_t *f = zip_fopen(z, "word/document.xml", 0);
    std::string doc_content(st.size, '\0');
    zip_fread(f, doc_content.data(), st.size);
    zip_fclose(f);

    REQUIRE(doc_content.find("<c:chart") != std::string::npos);
    // Verify EMUs calculated from twips (6000 * 635 = 3810000)
    REQUIRE(doc_content.find("cx=\"3810000\"") != std::string::npos);
    // Height: 4000 * 635 = 2540000
    REQUIRE(doc_content.find("cy=\"2540000\"") != std::string::npos);

    // Verify the second chart (Bar) has correct axes inserted
    REQUIRE(zip_stat(z, "word/charts/chart2.xml", 0, &st) == 0);
    zip_file_t *f_chart2 = zip_fopen(z, "word/charts/chart2.xml", 0);
    std::string chart2_content(st.size, '\0');
    zip_fread(f_chart2, chart2_content.data(), st.size);
    zip_fclose(f_chart2);

    REQUIRE(chart2_content.find("<c:barChart>") != std::string::npos);
    REQUIRE(chart2_content.find("<c:catAx>") != std::string::npos); // category axis
    REQUIRE(chart2_content.find("<c:valAx>") != std::string::npos); // value axis

    zip_close(z);
}

TEST_CASE("Chart Validator testing", "[chart][validator]") {
    const char *xsd_path = "dummy_chart_schema.xsd";
    {
        std::ofstream out(xsd_path);
        out << "<?xml version=\"1.0\"?>"
            << "<xs:schema xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" "
            << "targetNamespace=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
            << "xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
            << "elementFormDefault=\"qualified\">"
            << "  <xs:element name=\"chartSpace\">"
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
    ChartSeries s{"Data", {"A"}, {}, {1.0}, ""};
    doc.addChart(ChartType::Line, {s});

    std::string errs;
    bool result = doc.validate("word/charts/chart1.xml", validator, errs);

    // Note: pugixml serialization + simple validator typically passes or fails based on root node matching.
    // Here we mainly test that Document::validate routes correctly to the internal chart XML buffer.
    REQUIRE(result == true);
    REQUIRE(errs.empty());

    bool res2 = doc.validate("word/charts/chart99.xml", validator, errs);
    REQUIRE(res2 == false);
    REQUIRE(errs.find("Chart not found") != std::string::npos);

    std::remove(xsd_path);
}