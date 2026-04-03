#include <catch2/catch_test_macros.hpp>
#include <openword/Document.h>

using namespace openword;

TEST_CASE("Template Engine: cloneRowAndSetValues", "[template][table]") {
    Document doc;
    auto table = doc.addTable(2, 2);
    table.cell(0, 0).paragraphs()[0].addRun("Header 1");
    table.cell(0, 1).paragraphs()[0].addRun("Header 2");

    // Template row
    table.cell(1, 0).paragraphs()[0].addRun("${name}");
    table.cell(1, 1).paragraphs()[0].addRun("${age}");

    std::vector<std::map<std::string, std::string>> data = {{{"${name}", "Alice"}, {"${age}", "30"}},
                                                            {{"${name}", "Bob"}, {"${age}", "25"}},
                                                            {{"${name}", "Charlie"}, {"${age}", "40"}}};

    int count = doc.cloneRowAndSetValues("${name}", data);
    REQUIRE(count == 3);

    auto rows = table.rows();
    // 1 header + 3 cloned rows
    REQUIRE(rows.size() == 4);

    REQUIRE(table.cell(1, 0).text() == "Alice");
    REQUIRE(table.cell(1, 1).text() == "30");

    REQUIRE(table.cell(2, 0).text() == "Bob");
    REQUIRE(table.cell(2, 1).text() == "25");

    REQUIRE(table.cell(3, 0).text() == "Charlie");
    REQUIRE(table.cell(3, 1).text() == "40");
}
