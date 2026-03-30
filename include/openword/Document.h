#pragma once

#include <string>
#include <memory>
#include <vector>
#include <map>
#include <variant>
#include <gsl/gsl>

namespace openword {

class Paragraph;
class Table;

enum class TOCLeader {
    Dot,
    Hyphen,
    Underscore,
    None
};

class Section;

using BlockElement = std::variant<Paragraph, Table, Section>;


enum class HighlightColor {
    Yellow, Green, Cyan, Magenta, Blue, Red, DarkBlue, DarkCyan, DarkGreen, DarkMagenta, DarkRed, DarkYellow, DarkGray, LightGray, Black, Default
};

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
    int top = 1440;
    int right = 1440;
    int bottom = 1440;
    int left = 1440;
    int header = 720;
    int footer = 720;
    int gutter = 0;
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

struct CustomProperty {
    enum class Type { Text, Integer, Boolean, Double, Date };
    Type type;
    std::string value; // String representation
};

struct Metadata {
    // --- Core Properties (Summary tab) ---
    std::string title;
    std::string subject;
    std::string author; // Creator
    std::string keywords;
    std::string comments; // Description
    std::string lastModifiedBy;
    std::string category;
    
    // --- Extended Properties (Summary/Content tab) ---
    std::string company;
    std::string manager;
    std::string hyperlinkBase;
    
    // --- Custom Properties (Custom tab) ---
    std::map<std::string, CustomProperty> customProperties;
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
    Row& setRepeatHeaderRow(bool repeat = true);
    Row& setCantSplit(bool cantSplit = true);
    
    // --- Data Extraction ---
    std::vector<Cell> cells() const;
    std::string text() const;
    
    // --- DOM Mutability ---
    void remove();
    Row cloneAfter();
    int replaceText(const std::string& search, const std::string& replace);

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
    Run& setHighlight(HighlightColor color);
    Run& setCharacterSpacing(int twips);
    Run& addComment(int commentId);
    
    // --- Backgrounds ---
    Run& setHighlight(gsl::czstring color); // e.g. "yellow", "cyan"
    Run& setShading(const Color& fillColor);
    
    // --- Data Extractors ---
    bool isBold() const;
    bool isItalic() const;
    bool isStrike() const;
    HighlightColor highlightColor() const;
    std::string text() const;

private:
    void* node_; 
};

/**
 * @brief Represents a Paragraph (w:p).
 * This is a lightweight proxy object passed by value.
 */
class TextBox;

class Paragraph {
public:
    explicit Paragraph(void* node); // Internal use
    
    Run addRun(const std::string& text = "");
    void addRawXml(const std::string& xml);
    
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
    Paragraph& setList(int numId, int level = 0);
    
    // --- Borders & Shading ---
    Paragraph& setBorders(const BorderSettings& top, const BorderSettings& bottom, const BorderSettings& left, const BorderSettings& right);
    Paragraph& setBorders(const BorderSettings& all);
    Paragraph& setShading(const std::string& hexColor);
    
    // --- Text Box ---
    TextBox addTextBox(long long widthEMU, long long heightEMU, long long xOffsetEMU = 0, long long yOffsetEMU = 0);
    
    Run addHyperlink(gsl::czstring text, gsl::czstring url);
    Run addInternalLink(gsl::czstring text, gsl::czstring bookmarkName);
    void insertBookmark(gsl::czstring name);
    void addFootnoteReference(int footnoteId);
    void addEndnoteReference(int endnoteId);

    Section appendSectionBreak();

    // --- DOM Traversal & Data Extractors ---
    std::string styleId() const;
    std::string alignment() const;
    bool isList() const;
    int listLevel() const;
    std::vector<Run> runs() const;
    std::string text() const;
    int replaceText(const std::string& search, const std::string& replace);
    
    // --- DOM Mutability ---
    void remove();
    Paragraph cloneAfter();
    Paragraph insertParagraphAfter(const std::string& text = "");
    Table insertTableAfter(int rows, int cols);

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
    int gridSpan() const;
    std::string vMerge() const;
    std::vector<BlockElement> elements() const;
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
    int gridSpan() const;
    std::string vMerge() const;
    std::vector<BlockElement> elements() const;
    int replaceText(const std::string& search, const std::string& replace);
    std::string text() const;
    
    // --- DOM Mutability ---
    void remove();
    Table cloneAfter();
    Paragraph insertParagraphAfter(const std::string& text = "");

private:
    void* node_;
};

/**
 * @brief Represents a Microsoft Word Document (.docx)
 */

enum class NumberingFormat {
    Decimal,
    LowerLetter,
    UpperLetter,
    LowerRoman,
    UpperRoman,
    Bullet
};

struct ListLevel {
    int levelIndex = 0; // 0 to 8
    NumberingFormat format = NumberingFormat::Decimal;
    std::string text; // e.g. "%1.", "%1.%2.", or bullet char
    std::string jc = "left"; // Justification: left, center, right
    int indentTwips = 720;
    int hangingTwips = 360;
    std::string fontAscii = ""; // Specific font for bullets, e.g. "Symbol"
};

class AbstractNumbering {
public:
    explicit AbstractNumbering(void* node);
    AbstractNumbering& addLevel(const ListLevel& level);
private:
    void* node_;
};

class NumberingCollection {
public:
    explicit NumberingCollection(void* node);
    AbstractNumbering addAbstractNumbering(int abstractNumId);
    int addList(int abstractNumId, int restartNumId = -1);
private:
    void* node_;
};

class TextBox {
public:
    explicit TextBox(void* node);
    
    // TextBox acts as a miniature container
    Paragraph addParagraph(const std::string& text = "");
    
    TextBox& setFillColor(const std::string& hexColor);
    TextBox& setLineColor(const std::string& hexColor);
    
private:
    void* node_;
};

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

    void setEvenAndOddHeaders(bool val = true);

    // --- Content Creation ---
    Paragraph addParagraph(const std::string& text = "");
    Table addTable(int rows, int cols);
    
    // --- DOM Traversal & Extraction ---
    std::vector<BlockElement> elements() const;
    Section finalSection();
    std::vector<Paragraph> paragraphs() const;
    std::vector<Table> tables() const;
    
    void setMetadata(const Metadata& meta);
    Metadata metadata() const;

    void addTableOfContents(gsl::czstring title = "Table of Contents", int max_levels = 3, TOCLeader leader = TOCLeader::Dot);
    void addWatermark(const std::string& text);
    
    int createFootnote(const std::string& text);
    int createComment(const std::string& text, const std::string& author = "Author", const std::string& initials = "");
    int createEndnote(const std::string& text);
    
    StyleCollection styles();
    NumberingCollection numbering();
    
    /**

     * @param styleId Unique ID for the style (e.g., "Heading1")
     * @param name Human-readable name
     */
    void addStyle(gsl::czstring styleId, gsl::czstring name);

    // --- Utilities ---
    std::string convertMathMLToOMML(const std::string& mathml) const;
    std::string convertLaTeXToOMML(const std::string& latex) const;
    int replaceText(const std::string& search, const std::string& replace);
    
    // --- DOM Mutability ---
    void remove();
    Paragraph cloneAfter();
    Paragraph insertParagraphAfter(const std::string& text = "");
    Table insertTableAfter(int rows, int cols);

private:
    struct Impl;
    std::unique_ptr<Impl> pimpl;
};

} // namespace openword
