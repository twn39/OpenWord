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

enum class Orientation {
    Portrait,
    Landscape
};

struct Margins {
    uint32_t top = 1440;
    uint32_t right = 1440;
    uint32_t bottom = 1440;
    uint32_t left = 1440;
    uint32_t header = 720;
    uint32_t footer = 720;
    uint32_t gutter = 0;
};

enum class ListType {
    None,
    Bullet,
    Numbered
};

enum class HeaderFooterType {
    Default,
    First,
    Even
};

struct Metadata {
    std::string title;
    std::string author;
    std::string subject;
    std::string company;
    std::string creation_time; // ISO8601
};

class Paragraph;
class Header {
public:
    explicit Header(void* node);
    Paragraph addParagraph(const std::string& text = "");
private:
    void* node_;
};

class Footer {
public:
    explicit Footer(void* node);
    Paragraph addParagraph(const std::string& text = "");
private:
    void* node_;
};

class Section {
public:
    explicit Section(void* node);
    Section& setPageSize(uint32_t w_twips, uint32_t h_twips, Orientation orient = Orientation::Portrait);
    Section& setMargins(const Margins& margins);
    
    Header addHeader(HeaderFooterType type = HeaderFooterType::Default);
    Footer addFooter(HeaderFooterType type = HeaderFooterType::Default);
    Section& setColumns(int count, int spaceTwips = 720);
private:
    void* node_;
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
    
    /**
     * @brief Adds a raw OMML (Office Math Markup Language) equation to the paragraph.
     * @param omml The raw OMML XML string.
     */
    void addEquation(const std::string& omml);

    Paragraph& setStyle(gsl::czstring styleId);
    Paragraph& setOutlineLevel(int level); // 0 = Level 1 (Heading 1), 1 = Level 2, etc.
    Paragraph& setAlignment(gsl::czstring align); // "left", "center", "right", "both"
    Paragraph& setSpacing(int beforeTwips, int afterTwips, int lineSpacing = -1, gsl::czstring lineRule = "auto");
    Paragraph& setIndentation(int leftTwips, int rightTwips, int firstLineTwips = 0, int hangingTwips = 0);
    Paragraph& setList(ListType type, int level = 0);
    
    Run addHyperlink(gsl::czstring text, gsl::czstring url);
    Run addInternalLink(gsl::czstring text, gsl::czstring bookmarkName);
    void insertBookmark(gsl::czstring name);
    void addFootnoteReference(int footnoteId);
    void addEndnoteReference(int endnoteId);

    Section appendSectionBreak();

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
    Section finalSection();
    std::vector<Paragraph> paragraphs() const;
    
    void setMetadata(const Metadata& meta);
    Metadata metadata() const;

    void addTableOfContents(gsl::czstring title = "Table of Contents", int max_levels = 3);
    
    int createFootnote(const std::string& text);
    int createEndnote(const std::string& text);
    
    /**
     * @brief Define a basic paragraph style
     * @param styleId Unique ID for the style (e.g., "Heading1")
     * @param name Human-readable name
     */
    void addStyle(gsl::czstring styleId, gsl::czstring name);

    // --- Utilities ---
    std::string convertMathMLToOMML(const std::string& mathml) const;
    std::string convertLaTeXToOMML(const std::string& latex) const;

private:
    struct Impl;
    std::unique_ptr<Impl> pimpl;
};

} // namespace openword
