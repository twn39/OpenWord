#include <catch2/catch_test_macros.hpp>
#include <openword/Document.h>
#include <zip.h>

using namespace openword;

TEST_CASE("Chart Creation and XML Structure", "[chart]") {
    Document doc;
    ChartSeries s1{"Data 1", {"A", "B"}, {1.0, 2.0}, "FF0000"};

    ChartOptions options;
    options.title = "Test Chart Title";
    options.showDataLabels = true;
    options.legendPos = LegendPosition::Left;
    options.widthTwips = 6000;
    options.heightTwips = 4000;

    auto chart = doc.addChart(ChartType::Pie, {s1}, options);

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

    // Verify injected style data
    REQUIRE(chart_content.find("Test Chart Title") != std::string::npos);
    REQUIRE(chart_content.find("FF0000") != std::string::npos);
    REQUIRE(chart_content.find("<c:showVal val=\"1\"") != std::string::npos); // showDataLabels
    REQUIRE(chart_content.find("<c:legendPos val=\"l\"") != std::string::npos);

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

    zip_close(z);
}