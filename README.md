<div align="center">
  <h1>📖 OpenWord</h1>
  <p><strong>A modern, high-performance C++17 library for creating, parsing, and manipulating Microsoft Word (`.docx`) files.</strong></p>
  
  [![CI](https://github.com/twn39/OpenWord/actions/workflows/ci.yml/badge.svg)](https://github.com/twn39/OpenWord/actions/workflows/ci.yml)
  [![C++17](https://img.shields.io/badge/standard-C%2B%2B17-blue.svg)](https://en.cppreference.com/w/cpp/17)
  [![Rust](https://img.shields.io/badge/engine-Rust-orange.svg)](https://www.rust-lang.org/)
  [![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)
</div>

---

**OpenWord** stands out by offering a unique **hybrid Rust/C++ math engine**, providing seamless out-of-the-box support for translating **LaTeX** and **MathML** directly into native Word equations (OMML). 

It is built with a zero-compromise approach to performance and safety, strictly adhering to the **C++ Core Guidelines** (Microsoft GSL) and utilizing a fluent, ergonomic API design.

## ✨ Key Features

- 🧮 **Native Equations**: Convert LaTeX and MathML directly to Word OMML using a blazing-fast Rust bridge.
- 🎨 **Rich Formatting**: Fluent API chaining for text runs, custom colors, spacing, line breaks, and highlights.
- 📊 **Advanced Tables**: Full support for nested tables, cell merging (rowspan/colspan), borders, and shading.
- 📑 **Layout & Structure**: Multi-section documents, independent headers/footers, dynamic page numbers, and custom margins.
- 📝 **Lists & Styles**: One-click bullet/numbered list generation and comprehensive style inheritance trees (`docDefaults`).
- 💬 **Annotations**: Built-in support for document comments, footnotes, endnotes, and metadata injection.
- 📦 **Zero Deployment Friction**: XSLT stylesheets and math engine assets are statically embedded directly into the binary.

## 📚 Documentation

- **[Tutorial & Examples](docs/tutorial.md)**: A comprehensive, step-by-step developer guide with code snippets.
- **API Reference**: Generate local HTML documentation using Doxygen (`make doc_doxygen`).

## 🛠 Technical Stack

OpenWord automatically manages its C++ dependencies via CMake `FetchContent`. No manual downloads required.

- **C++17**: Core library architecture.
- **Rust (2021)**: LaTeX to MathML conversion (via `tex2math`).
- **libxml2 & libxslt**: Lightning-fast XML manipulation and MathML-to-OMML XSLT transformations.
- **pugixml**: Zero-allocation DOM parsing.
- **libzip**: Secure `.docx` archive packing.
- **Catch2**: Extensive automated test matrix.

## 🚀 Quick Start

### Prerequisites
- **CMake 3.15+**
- **C++17 Compiler** (GCC 9+, Clang 10+, Apple Clang, or MSVC 2019+)
- **Rust Toolchain** (`cargo`)

### Build Instructions

**1. Compile the Rust Math Engine**
```bash
cd rust_math_engine
cargo build --release
cd ..
```

**2. Configure and Build the C++ Project**
```bash
cmake -S . -B build -DCMAKE_BUILD_TYPE=Release
cmake --build build --config Release -j 4
```

**3. Run the Test Suite & Demo**
```bash
cd build
ctest --output-on-failure -C Release
./openword_demo
```

## 💻 Usage Example

OpenWord features a highly ergonomic, modern C++ API designed to reduce boilerplate.

```cpp
#include <openword/Document.h>

int main() {
    openword::Document doc;

    // 1. Fluent Text Formatting
    auto p = doc.addParagraph("Welcome to ");
    p.addRun("OpenWord").setBold(true).setColor(openword::Color(0, 128, 255));
    
    // 2. Hybrid Rust Math Engine (LaTeX to Native Word Equation)
    std::string omml = doc.convertLaTeXToOMML("\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}");
    doc.addParagraph().addEquation(omml);

    // 3. Ergonomic Tables
    auto t = doc.addTable(2, 2);
    t.setBorders(openword::BorderSettings{openword::BorderStyle::Dashed, 4, "000000"});
    t.cell(0, 0).addParagraph("Top Left");
    t.cell(1, 1).addParagraph("Nested tables are supported too!");

    doc.save("demo_output.docx");
    return 0;
}
```

## ⚖️ License
OpenWord is released under the **MIT License**. Reference libraries located in `third_party/` belong to their respective owners.
