# OpenWord

**OpenWord** is a modern, high-performance C++17 library designed for creating, parsing, and manipulating Microsoft Word (`.docx`) files. It features a unique **hybrid Rust/C++ math engine** that provides seamless support for LaTeX and MathML equations.

## Key Features

- **Equations Support**: Convert LaTeX and MathML to native Word equations (OMML) using a high-performance Rust bridge.
- **Modern C++ Architecture**: Built with C++17, following strict safety standards using **Microsoft GSL**.
- **Self-Contained Build**: Automatically fetches and builds all C++ dependencies via CMake `FetchContent`.
- **Comprehensive API**: Easy-to-use interface for paragraphs, runs, tables, images, and complex multi-section documents.
- **Styling & Formatting**: Full control over fonts, colors, spacing, indentation, and list numbering.

## Technical Stack

- **C++17**: Core library implementation.
- **Rust (2021 Edition)**: LaTeX to MathML conversion (via `latex2mathml`).
- **libxslt & libxml2**: MathML to OMML (Word Math) transformations.
- **pugixml**: High-speed XML parsing and DOM manipulation.
- **libzip**: Secure `.docx` (ZIP) archive management.
- **fmt**: Modern string formatting.
- **Catch2**: Robust unit and integration testing.

## Prerequisites

- **CMake 3.15+**
- **C++17 Compliant Compiler** (GCC 9+, Clang 10+, or MSVC 2019+)
- **Rust Toolchain** (Cargo) for the math engine.

## Quick Start

### 1. Build the Rust Math Engine
```bash
cd rust_math_engine
cargo build --release
cd ..
```

### 2. Build the C++ Project
```bash
cmake -S . -B build
cmake --build build
```

### 3. Run Demo & Tests
```bash
./build/openword_demo
cd build && ctest --output-on-failure
```

## Usage Example

```cpp
#include <openword/Document.h>

int main() {
    openword::Document doc;

    // Add a styled heading
    doc.addStyle("MyHeading", "Heading 1");
    auto p1 = doc.addParagraph("OpenWord Math Demo");
    p1.setStyle("MyHeading").setAlignment("center");

    // Add a paragraph with a LaTeX equation
    auto p2 = doc.addParagraph("Behold the power of Rust + C++: ");
    std::string omml = doc.convertLaTeXToOMML("\\frac{-b \pm \sqrt{b^2 - 4ac}}{2a}");
    p2.addEquation(omml);

    // Save the document
    doc.save("math_demo.docx");
    return 0;
}
```

## License
OpenWord is released under the MIT License. Reference libraries in `third_party/` belong to their respective owners.
