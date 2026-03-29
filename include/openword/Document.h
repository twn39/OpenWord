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

enum class BorderStyle { None, Single, Dashed, Dotted, Double, Thick };
struct BorderSettings {
    BorderStyle style = BorderStyle::Single;
    int size = 4; // 1/8 points
    std::string color = "auto";
};

enum class VerticalAlignment { Top, Center, Bottom };
enum class HeightRule { Auto, AtLeast, Exact };

struct Metadata {
    std::string title;
    std::string author;
    std::string subject;
    std::string company;
    std::string creation_time; // ISO8601
};


class Font {
public:
    explicit Font(void* node);
    Font& setSize(int halfPoints);
    Font& setBold(bool val = true);
    Font& setItalic(bool val = true);
    Font& setColor(const std::string& hexColor);
    Font& setName(const std::string& ascii);
private:
    void* node_;
};

class ParagraphFormat {
public:
    explicit ParagraphFormat(void* node);
    ParagraphFormat& setOutlineLevel(int level);
    ParagraphFormat& setSpacing(int beforeTwips, int afterTwips);
private:
    void* node_;
};

class Style {
public:
    explicit Style(void* node);
    Font getFont();
    ParagraphFormat getParagraphFormat();
    Style& setName(const std::string& name);
private:
    void* node_;
};

class StyleCollection {
public:
    explicit StyleCollection(void* node);
    Style get(const std::string& styleId);
    Style add(const std::string& styleId, const std::string& type = "paragraph");
private:
    void* node_;
};


enum class ImagePosition {
    Inline,       // In line with text
    BehindText,   // Floating behind text
    InFrontOfText // Floating in front of text
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
class Row {
public:
    explicit Row(void* node);
    Row& setHeight(int twips, HeightRule rule = HeightRule::AtLeast);
    
    // --- Data Extraction ---
    std::vector<Cell> cells() const;
    std::string text() const;
private:
    void* node_;
};

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
     * @param position Defines how the image is positioned relative to text.
     * @param xOffset Optional horizontal offset in EMU (English Metric Units). Only applicable if not Inline.
     * @param yOffset Optional vertical offset in EMU (English Metric Units). Only applicable if not Inline.
     */
    void addImage(gsl::czstring image_path, double scale = 1.0, ImagePosition position = ImagePosition::Inline, long long xOffset = 0, long long yOffset = 0);
    
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
    int replaceText(const std::string& search, const std::string& replace);

private:
    void* node_;
};

/**
 * @brief Represents a cell in a Table in the Word document.
 */
class Table;
class Cell {
public:
    explicit Cell(void* node);

    // --- Content Creation ---
    Paragraph addParagraph(const std::string& text = "");

    Cell& setVertAlign(VerticalAlignment align);
    Cell& setShading(const std::string& hexColor);
    Cell& setWidth(int twips, const std::string& type = "dxa");
    Cell& setBorders(const BorderSettings& top, const BorderSettings& bottom, const BorderSettings& left, const BorderSettings& right);
    Cell& setBorders(const BorderSettings& all);
    
    // --- Data Extraction & Manipulation ---
    int replaceText(const std::string& search, const std::string& replace);
    std::string text() const;
    std::vector<Paragraph> paragraphs() const;

private:
    void* node_;
};

/**
 * @brief Represents a Table in the Word document.
 */
class Table {
public:
    explicit Table(void* node);

    // --- Data Access & Traversal ---
    Cell cell(int row, int col);
    Row row(int rowIndex);
    std::vector<Row> rows() const;
    void mergeCells(int startRow, int startCol, int endRow, int endCol);
    
    // --- Ergonomic Table Formatting ---
    Table& setBorders(const BorderSettings& all);
    Table& setBorders(const BorderSettings& outer, const BorderSettings& inner);
    Table& setBorders(const BorderSettings& top, const BorderSettings& bottom, const BorderSettings& left, const BorderSettings& right, const BorderSettings& insideH, const BorderSettings& insideV);
    
    Table& setColumnWidth(int colIndex, int twips);
    Table& setColumnWidths(const std::vector<int>& twipsList);
    
    Table& setAlignment(gsl::czstring align);
    
    // --- Data Extraction & Manipulation ---
    int replaceText(const std::string& search, const std::string& replace);
    std::string text() const;

private:
    void* node_;
};

/**
 * @brief Represents a Microsoft Word Document (.docx)
 */
class Document {
public:
    Document();
    explicit Document(const std::string& templatePath);
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
    
    // --- DOM Traversal & Extraction ---
    Section finalSection();
    std::vector<Paragraph> paragraphs() const;
    std::vector<Table> tables() const;
    
    void setMetadata(const Metadata& meta);
    Metadata metadata() const;

    void addTableOfContents(gsl::czstring title = "Table of Contents", int max_levels = 3);
    
    int createFootnote(const std::string& text);
    int createEndnote(const std::string& text);
    
    StyleCollection styles();
    
    /**

     * @param styleId Unique ID for the style (e.g., "Heading1")
     * @param name Human-readable name
     */
    void addStyle(gsl::czstring styleId, gsl::czstring name);

    // --- Utilities ---
    std::string convertMathMLToOMML(const std::string& mathml) const;
    std::string convertLaTeXToOMML(const std::string& latex) const;
    int replaceText(const std::string& search, const std::string& replace);

private:
    struct Impl;
    std::unique_ptr<Impl> pimpl;
};

} // namespace openword
