#include <openword/Document.h>
#include <fmt/core.h>
#include <vector>
#include <string>

void test_basic_text() {
    openword::Document doc;
    doc.addParagraph("This is a basic text test.");
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
    p.addRun("Big. ").setFontSize(48); 
    p.addRun("Colored. ").setColor(openword::Color(0, 128, 255));

    auto p2 = doc.addParagraph("This paragraph has custom indentation and spacing.");
    p2.setIndentation(360, 0, 720);
    p2.setSpacing(240, 120);

    doc.save("test_02_formatting.docx");
    fmt::print("- test_02_formatting.docx (Text Formatting)\n");
}

void test_styles() {
    openword::Document doc;
    doc.addStyle("MyHeading1", "Heading 1");
    auto p1 = doc.addParagraph("Styled Heading");
    p1.setStyle("MyHeading1").setAlignment("center");
    doc.save("test_03_styles.docx");
    fmt::print("- test_03_styles.docx (Styles)\n");
}

void test_image() {
    openword::Document doc;
    doc.addParagraph("Images test:");
    doc.addParagraph().addImage("tests/test.jpg", 0.5);
    doc.save("test_04_image.docx");
    fmt::print("- test_04_image.docx (Images)\n");
}

void test_tables() {
    openword::Document doc;
    doc.addParagraph("Table test:");
    auto table = doc.addTable(3, 3);
    table.mergeCells(0, 0, 0, 2);
    table.cell(0, 0).addParagraph("Merged Header");
    doc.save("test_05_tables.docx");
    fmt::print("- test_05_tables.docx (Tables)\n");
}

void test_sections_and_headers() {
    openword::Document doc;
    doc.finalSection().setPageSize(16838, 11906, openword::Orientation::Landscape);
    doc.addParagraph("Landscape content");
    doc.save("test_06_sections.docx");
    fmt::print("- test_06_sections.docx (Sections)\n");
}

void test_lists() {
    openword::Document doc;
    doc.addParagraph("List test:");
    doc.addParagraph("Item 1").setList(openword::ListType::Bullet, 0);
    doc.save("test_07_lists.docx");
    fmt::print("- test_07_lists.docx (Lists)\n");
}

void test_math() {
    openword::Document doc;
    doc.addParagraph("LaTeX Comprehensive Capability Test").setStyle("Heading1");

    auto add_latex = [&](const std::string& label, const std::string& latex) {
        doc.addParagraph(label + " (" + latex + "):");
        std::string omml = doc.convertLaTeXToOMML(latex);
        if (!omml.empty()) {
            doc.addParagraph().addEquation(omml);
        } else {
            fmt::print("- Critical failure (empty output) for: {}\n", latex);
        }
    };

    // 1. Symbol Varieties
    add_latex("Greek Variants", "\\varepsilon \\varkappa \\varpi \\varrho \\varsigma \\varphi");
    add_latex("Hebrew Symbols", "\\aleph \\beth \\gimel \\daleth");
    add_latex("Logic Symbols", "\\forall \\exists \\neg \\lnot \\wedge \\vee \\Rightarrow \\iff");

    // 2. Operators & Limits
    add_latex("Big Operators", "\\sum_{i=1}^n \\prod_{j=1}^m \\coprod \\oint_\\gamma \\bigcap \\bigcup \\bigoplus \\bigotimes");
    add_latex("Integrals", "\\int_a^b x^2 dx \\quad \\iint_D f(x,y) dS \\quad \\iiint_V g(x,y,z) dV");
    add_latex("Limits", "\\lim_{x \\to \\infty} \\frac{1}{x} = 0");

    // 3. Delimiters & Accents
    add_latex("Large Brackets", "\\left( \\sum_{k=1}^n \\frac{1}{k} \\right) \\quad \\left[ \\int \\right] \\quad \\left\\{ \\dots \\right\\}");
    add_latex("Accents", "\\hat{a} \\check{b} \\tilde{c} \\acute{d} \\grave{e} \\dot{f} \\ddot{g} \\breve{h} \\bar{i} \\vec{j}");

    // 4. Advanced Structures
    add_latex("Matrices (Standard)", "\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix} \\quad \\begin{bmatrix} 1 & 0 \\\\ 0 & 1 \\end{bmatrix}");
    add_latex("Matrices (Advanced)", "\\begin{vmatrix} x & y \\\\ z & w \\end{vmatrix} \\quad \\begin{Bmatrix} p & q \\\\ r & s \\end{Bmatrix}");
    add_latex("Cases / Conditionals", "f(x) = \\begin{cases} x & x > 0 \\\\ 0 & x \\leq 0 \\end{cases}");
    add_latex("Continued Fractions", "\\phi = 1 + \\frac{1}{1 + \\frac{1}{1 + \\dots}}");

    // 5. Physics & Field Equations
    add_latex("Maxwell", "\\nabla \\cdot \\mathbf{E} = \\frac{\\rho}{\\varepsilon_0} \\quad \\nabla \\times \\mathbf{B} - \\mu_0\\varepsilon_0\\frac{\\partial\\mathbf{E}}{\\partial t} = \\mu_0\\mathbf{J}");
    add_latex("Schrodinger", "i\\hbar\\frac{\\partial}{\\partial t}\\Psi(\\mathbf{r},t) = \\hat H \\Psi(\\mathbf{r},t)");

    // 6. Over/Under decorations
    add_latex("Over/Under", "\\overline{z} \\quad \\underline{u} \\quad \\overbrace{a+b+c}^{3} \\quad \\underbrace{x+y+z}_{n}");

    // 7. Set Theory & Logic Advanced
    add_latex("Set Relations", "A \\subsetneq B \\quad C \\parallel D \\quad x \\perp y \\quad \\complement_A B");
    add_latex("Logic Chains", "P \\implies Q \\iff \\neg Q \\implies \\neg P \\quad \\exists! x : P(x)");

    // 8. Number Theory & Arithmetic
    add_latex("Modulo & Divisibility", "a \\equiv b \\pmod{n} \\quad d \\mid n \\quad \\gcd(a,b) = 1");
    add_latex("Binary Ops", "a \\oplus b \\quad c \\otimes d \\quad x \\odot y \\quad \\ast \\star \\diamond");

    // 9. Advanced Calculus
    add_latex("Surface Integral", "\\oiint_S \\mathbf{E} \\cdot d\\mathbf{A} = \\frac{Q_{enc}}{\\varepsilon_0}");
    add_latex("N-th Derivative", "\\frac{d^n y}{dx^n} \\quad f^{(n)}(x) \\quad \\dot{x} \\quad \\ddot{x}");

    // 10. Symbols Extravaganza
    add_latex("Hebrew & Misc", "\\aleph_0 \\beth \\gimel \\daleth \\quad \\ell \\wp \\Re \\Im \\hbar \\mho");
    add_latex("Relation Symbols", "= \\neq < > \\leq \\geq \\approx \\sim \\equiv \\subset \\subseteq \\in");

    // 11. Large Multi-line (Simulated via simple grouping)
    add_latex("Multi-line Structure", "\\left( \\sum_{i=1}^n x_i \\right) \\left( \\prod_{j=1}^m y_j \\right) = \\int \\dots \\int f(x_1, \\dots, x_k) \\, dV");

    doc.save("test_08_math.docx");
    fmt::print("- test_08_math.docx (LaTeX Full Capability Showcase Generated)\n");
}

int main() {
    fmt::print("Generating capability test files...\n");
    test_basic_text();
    test_text_formatting();
    test_styles();
    test_image();
    test_tables();
    test_sections_and_headers();
    test_lists();
    test_math();
    fmt::print("\nDone! Please verify test_08_math.docx to see the engine's limits.\n");
    return 0;
}
