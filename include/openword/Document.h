#pragma once

#include <string>
#include <memory>
#include <vector>
#include <gsl/gsl>

namespace openword {

struct Color {
    uint8_t r{0};
    uint8_t g{0};
    uint8_t b{0};
    bool is_auto{true};

    Color() = default;
    Color(uint8_t red, uint8_t green, uint8_t blue) : r(red), g(green), b(blue), is_auto(false) {}
    
    static Color Auto() { return Color(); }
    std::string hex() const;
};

enum class VertAlign {
    Baseline,
    Superscript,
    Subscript
};

/**
 * @brief Represents a Run of text within a paragraph (w:r).
 * This is a lightweight proxy object passed by value.
 */
class Cell;
class Table;
class Run {
public:
    explicit Run(void* node); // Internal use

    Run& setText(const std::string& text);
    
    // --- Font Properties ---
    Run& setFontFamily(gsl::czstring ascii, gsl::czstring eastAsia = "");
    Run& setFontSize(int halfPoints);
    Run& setColor(const Color& color);
    
    // --- Text Styles ---
    Run& setBold(bool val = true);
    Run& setItalic(bool val = true);
    Run& setUnderline(gsl::czstring val = "single");
    Run& setStrike(bool val = true);
    Run& setDoubleStrike(bool val = true);
    Run& setVertAlign(VertAlign align);
    
    // --- Backgrounds ---
    Run& setHighlight(gsl::czstring color); // e.g. "yellow", "cyan"
    Run& setShading(const Color& fillColor);
    
    // --- Data Extractors ---
    std::string text() const;

private:
    void* node_; 
};

/**
 * @brief Represents a Paragraph (w:p).
 * This is a lightweight proxy object passed by value.
 */
class Paragraph {
public:
    explicit Paragraph(void* node); // Internal use
    
    Run addRun(const std::string& text = "");
    
    /**
     * @brief Adds an image to the paragraph.
     * @param image_path Path to the image file (JPEG or PNG).
     * @param scale Optional scaling factor (e.g., 0.5 for half size, 2.0 for double size). Default is 1.0.
     */
    void addImage(gsl::czstring image_path, double scale = 1.0);
    
    Paragraph& setStyle(gsl::czstring styleId);
    Paragraph& setAlignment(gsl::czstring align); // "left", "center", "right", "both"
    Paragraph& setSpacing(int beforeTwips, int afterTwips, int lineSpacing = -1, gsl::czstring lineRule = "auto");
    Paragraph& setIndentation(int leftTwips, int rightTwips, int firstLineTwips = 0, int hangingTwips = 0);

    // --- DOM Traversal & Data Extractors ---
    std::vector<Run> runs() const;
    std::string text() const;

private:
    void* node_;
};

/**
 * @brief Represents a cell in a Table in the Word document.
 */
class Cell {
public:
    explicit Cell(void* node);

    // --- Content Creation ---
    Paragraph addParagraph(const std::string& text = "");

private:
    void* node_;
};

/**
 * @brief Represents a Table in the Word document.
 */
class Table {
public:
    explicit Table(void* node);

    // --- Data Access ---
    Cell cell(int row, int col);
    void mergeCells(int startRow, int startCol, int endRow, int endCol);

private:
    void* node_;
};

/**
 * @brief Represents a Microsoft Word Document (.docx)
 */
class Document {
public:
    Document();
    ~Document();

    // Disable copy for safety (RAII)
    Document(const Document&) = delete;
    Document& operator=(const Document&) = delete;

    // --- IO Operations ---
    bool save(gsl::czstring filepath);
    bool load(gsl::czstring filepath);

    // --- Content Creation ---
    Paragraph addParagraph(const std::string& text = "");
    Table addTable(int rows, int cols);
    
    // --- DOM Traversal ---
    std::vector<Paragraph> paragraphs() const;
    
    /**
     * @brief Define a basic paragraph style
     * @param styleId Unique ID for the style (e.g., "Heading1")
     * @param name Human-readable name
     */
    void addStyle(gsl::czstring styleId, gsl::czstring name);

    // --- Utilities ---
    std::string convertMathMLToOMML(const std::string& mathml) const;

private:
    struct Impl;
    std::unique_ptr<Impl> pimpl;
};

} // namespace openword
