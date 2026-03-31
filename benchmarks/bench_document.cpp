#include <catch2/catch_test_macros.hpp>
#include <catch2/benchmark/catch_benchmark.hpp>
#include <openword/Document.h>

TEST_CASE("Document creation and population benchmarks", "[benchmark]") {
    BENCHMARK("Create empty Document") {
        return openword::Document();
    };

    BENCHMARK("Create Document and add 1000 paragraphs") {
        openword::Document doc;
        for (int i = 0; i < 1000; ++i) {
            doc.addParagraph("This is paragraph " + std::to_string(i));
        }
        return doc.paragraphs().size();
    };

    BENCHMARK("Create Document, add 100 paragraphs and save") {
        openword::Document doc;
        for (int i = 0; i < 100; ++i) {
            doc.addParagraph("This is paragraph " + std::to_string(i));
        }
        doc.save("bench_save.docx");
        return doc.paragraphs().size();
    };
    
    BENCHMARK("Create 10x10 Table") {
        openword::Document doc;
        auto table = doc.addTable(10, 10);
        for(size_t r = 0; r < 10; ++r) {
            for(size_t c = 0; c < 10; ++c) {
                table.cell(r, c).addParagraph("Cell " + std::to_string(r) + "-" + std::to_string(c));
            }
        }
        return doc.paragraphs().size();
    };
}

TEST_CASE("Math conversion benchmarks", "[benchmark]") {
    BENCHMARK("Convert Simple LaTeX") {
        openword::Document doc;
        std::string omml = doc.convertLaTeXToOMML("\\alpha + \\beta = \\gamma");
        doc.addParagraph("Math: ").addEquation(omml);
        return omml.size();
    };

    BENCHMARK("Convert Complex LaTeX (Schrodinger)") {
        openword::Document doc;
        std::string omml = doc.convertLaTeXToOMML("i\\hbar \\frac{\\partial}{\\partial t} \\Psi(\\mathbf{r}, t) = \\left[ -\\frac{\\hbar^2}{2m}\\nabla^2 + V(\\mathbf{r}, t) \\right] \\Psi(\\mathbf{r}, t)");
        doc.addParagraph("").addEquation(omml);
        return omml.size();
    };
}
