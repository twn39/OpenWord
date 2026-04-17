#pragma once

#include "Validator.h"

#include <gsl/gsl>
#include <map>
#include <memory>
#include <string>
#include <variant>
#include <vector>

namespace openword {

class Paragraph;
class Table;

enum class TOCLeader { Dot, Hyphen, Underscore, None };

class Section;

using BlockElement = std::variant<Paragraph, Table, Section>;

enum class HighlightColor {
    Yellow,
    Green,
    Cyan,
    Magenta,
    Blue,
    Red,
    DarkBlue,
    DarkCyan,
    DarkGreen,
    DarkMagenta,
    DarkRed,
    DarkYellow,
    DarkGray,
    LightGray,
    Black,
    Default
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

enum class VertAlign { Baseline, Superscript, Subscript };

enum class Orientation { Portrait, Landscape };

struct Margins {
    int top = 1440;
    int right = 1440;
    int bottom = 1440;
    int left = 1440;
    int header = 720;
    int footer = 720;
    int gutter = 0;
};

enum class ListType { None, Bullet, Numbered };

enum class HeaderFooterType { Default, First, Even };

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
    explicit Font(void *node);
    Font &setSize(int halfPoints);
    Font &setBold(bool val = true);
    Font &setItalic(bool val = true);
    Font &setColor(const std::string &hexColor);
    Font &setName(const std::string &ascii);

  private:
    void *node_;
};

class ParagraphFormat {
  public:
    explicit ParagraphFormat(void *node);
    ParagraphFormat &setOutlineLevel(int level);
    ParagraphFormat &setSpacing(int beforeTwips, int afterTwips);

  private:
    void *node_;
};

class Style {
  public:
    explicit Style(void *node);
    Font getFont();
    ParagraphFormat getParagraphFormat();

    // Core Identity
    Style &setName(const std::string &name);

    // Inheritance & Flow
    Style &setBasedOn(const std::string &parentStyleId);
    Style &setNextStyle(const std::string &nextStyleId);

    // UI Controls
    Style &setPrimary(bool isPrimary = true); // Shows up in Quick Styles gallery
    Style &setUiPriority(int priority);
    Style &setHidden(bool isHidden = true);

  private:
    void *node_;
};

class StyleCollection {
  public:
    explicit StyleCollection(void *node);

    // Default document properties (w:docDefaults)
    /**
     * @brief Retrieves the global document default font settings (`w:docDefaults/w:rPrDefault`).
     * Modifying this affects all text in the document lacking an explicit style.
     */
    Font getDefaultFont();

    /**
     * @brief Retrieves the global document default paragraph format settings (`w:docDefaults/w:pPrDefault`).
     */
    ParagraphFormat getDefaultParagraphFormat();

    Style get(const std::string &styleId);
    Style add(const std::string &styleId, const std::string &type = "paragraph");

  private:
    void *node_;
};

enum class ImagePosition {
    Inline,       // In line with text
    BehindText,   // Floating behind text
    InFrontOfText // Floating in front of text
};

class Paragraph;
class Header {
  public:
    explicit Header(void *node);
    Paragraph addParagraph(const std::string &text = "");
    void addHtml(const std::string &html);

  private:
    void *node_;
};

class Footer {
  public:
    explicit Footer(void *node);
    Paragraph addParagraph(const std::string &text = "");
    void addHtml(const std::string &html);

  private:
    void *node_;
};

class Section {
  public:
    explicit Section(void *node);
    Section &setPageSize(uint32_t w_twips, uint32_t h_twips, Orientation orient = Orientation::Portrait);
    Section &setMargins(const Margins &margins);

    Header addHeader(HeaderFooterType type = HeaderFooterType::Default);
    Footer addFooter(HeaderFooterType type = HeaderFooterType::Default);

    // --- Header/Footer Controls ---
    /**
     * @brief Removes the header of the specified type from this section and blocks inheritance from previous sections.
     */
    Section &removeHeader(HeaderFooterType type = HeaderFooterType::Default);

    /**
     * @brief Removes the footer of the specified type from this section and blocks inheritance.
     */
    Section &removeFooter(HeaderFooterType type = HeaderFooterType::Default);

    Section &setColumns(int count, int spaceTwips = 720);

  private:
    void *node_;
};

/**
 * @brief Represents a Run of text within a paragraph (w:r).
 * This is a lightweight proxy object passed by value.
 */
class Cell;
class Row {
  public:
    explicit Row(void *node);
    Row &setHeight(int twips, HeightRule rule = HeightRule::AtLeast);
    Row &setRepeatHeaderRow(bool repeat = true);
    Row &setCantSplit(bool cantSplit = true);

    // --- Data Extraction ---
    std::vector<Cell> cells() const;
    std::string text() const;

    // --- DOM Mutability ---
    void remove();
    Row cloneAfter();
    int replaceText(const std::string &search, const std::string &replace);

  private:
    void *node_;
};

class Cell;
class Table;
class Run {
  public:
    explicit Run(void *node); // Internal use

    Run &setText(const std::string &text);

    // --- Font Properties ---
    Run &setFontFamily(gsl::czstring ascii, gsl::czstring eastAsia = "");
    Run &setFontSize(int halfPoints);
    Run &setColor(const Color &color);

    // --- Text Styles ---
    Run &setBold(bool val = true);
    Run &setItalic(bool val = true);
    Run &setUnderline(gsl::czstring val = "single");
    Run &setStrike(bool val = true);
    Run &setDoubleStrike(bool val = true);
    Run &setVertAlign(VertAlign align);
    Run &setHighlight(HighlightColor color);
    Run &setCharacterSpacing(int twips);
    /**
     * @brief Anchors an existing comment to this specific text run.
     * @param commentId The unique ID of the comment returned by `Document::createComment`.
     * @return Reference to self for fluent chaining.
     */
    Run &addComment(int commentId);

    // --- Breaks ---
    /**
     * @brief Inserts a soft line break (`<w:br/>`) in the current run (equivalent to Shift+Enter).
     * @return Reference to self for fluent chaining.
     */
    Run &addLineBreak();

    /**
     * @brief Inserts a hard page break (`<w:br w:type="page"/>`) pushing subsequent text to the next page.
     * @return Reference to self for fluent chaining.
     */
    Run &addPageBreak();

    // --- Backgrounds ---
    Run &setHighlight(gsl::czstring color); // e.g. "yellow", "cyan"
    Run &setShading(const Color &fillColor);

    // --- Data Extractors ---
    bool isBold() const;
    bool isItalic() const;
    bool isStrike() const;
    HighlightColor highlightColor() const;
    std::string text() const;

    bool isFootnoteReference() const;
    int footnoteId() const;
    bool isEndnoteReference() const;
    int endnoteId() const;

  private:
    void *node_;
};

/**
 * @brief Represents a Paragraph (w:p).
 * This is a lightweight proxy object passed by value.
 */
class TextBox;

class Paragraph {
  public:
    explicit Paragraph(void *node); // Internal use

    Run addRun(const std::string &text = "");
    void addRawXml(const std::string &xml);
    /**
     * @brief Injects a dynamic `PAGE` field code to display the current page number.
     * @return The Run containing the field.
     */
    Run addPageNumber();

    /**
     * @brief Injects a dynamic `NUMPAGES` field code to display the total document page count.
     * @return The Run containing the field.
     */
    Run addTotalPages();

    /**
     * @brief Adds an image to the paragraph.
     * @param image_path Path to the image file (JPEG or PNG).
     * @param scale Optional scaling factor (e.g., 0.5 for half size, 2.0 for double size). Default is 1.0.
     * @param position Defines how the image is positioned relative to text.
     * @param xOffset Optional horizontal offset in EMU (English Metric Units). Only applicable if not Inline.
     * @param yOffset Optional vertical offset in EMU (English Metric Units). Only applicable if not Inline.
     */
    void addImage(gsl::czstring image_path, double scale = 1.0, ImagePosition position = ImagePosition::Inline,
                  long long xOffset = 0, long long yOffset = 0);

    /**
     * @brief Adds a raw OMML (Office Math Markup Language) equation to the paragraph.
     * @param omml The raw OMML XML string.
     */
    void addEquation(const std::string &omml);

    Paragraph &setStyle(gsl::czstring styleId);
    Paragraph &setOutlineLevel(int level);        // 0 = Level 1 (Heading 1), 1 = Level 2, etc.
    Paragraph &setAlignment(gsl::czstring align); // "left", "center", "right", "both"
    Paragraph &setSpacing(int beforeTwips, int afterTwips, int lineSpacing = -1, gsl::czstring lineRule = "auto");
    Paragraph &setIndentation(int leftTwips, int rightTwips, int firstLineTwips = 0, int hangingTwips = 0);
    Paragraph &setList(int numId, int level = 0);

    // --- Borders & Shading ---
    Paragraph &setBorders(const BorderSettings &top, const BorderSettings &bottom, const BorderSettings &left,
                          const BorderSettings &right);
    Paragraph &setBorders(const BorderSettings &all);
    Paragraph &setShading(const std::string &hexColor);

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
    int replaceText(const std::string &search, const std::string &replace);

    // --- DOM Mutability ---
    void remove();
    Paragraph cloneAfter();
    Paragraph insertParagraphAfter(const std::string &text = "");
    Table insertTableAfter(int rows, int cols);

  private:
    void *node_;
};

/**
 * @brief Represents a cell in a Table in the Word document.
 */
class Table;
class Cell {
  public:
    explicit Cell(void *node);

    // --- Content Creation ---
    Paragraph addParagraph(const std::string &text = "");
    void addHtml(const std::string &html);

    /**
     * @brief Nests a new table inside this cell.
     * @param rows Initial number of rows.
     * @param cols Initial number of columns.
     * @return The proxy object to the newly created nested table.
     */
    Table addTable(int rows, int cols);
    void addImage(gsl::czstring image_path, double scale = 1.0, ImagePosition position = ImagePosition::Inline,
                  long long xOffset = 0, long long yOffset = 0);

    Cell &setVertAlign(VerticalAlignment align);
    Cell &setShading(const std::string &hexColor);
    Cell &setWidth(int twips, const std::string &type = "dxa");
    Cell &setBorders(const BorderSettings &top, const BorderSettings &bottom, const BorderSettings &left,
                     const BorderSettings &right);
    Cell &setBorders(const BorderSettings &all);

    // --- Data Extraction & Manipulation ---
    int gridSpan() const;
    std::string vMerge() const;
    std::vector<BlockElement> elements() const;
    int replaceText(const std::string &search, const std::string &replace);
    std::string text() const;
    std::vector<Paragraph> paragraphs() const;

  private:
    void *node_;
};

/**
 * @brief Represents a Table in the Word document.
 *
 * @par Example: Creating and Formatting a Table
 * @code
 * auto t = doc.addTable(2, 2);
 * t.setBorders(openword::BorderSettings{openword::BorderStyle::Single, 4, "000000"});
 * t.cell(0, 0).addParagraph("Top Left");
 * t.row(0).setRepeatHeaderRow(true);
 * @endcode
 */
class Table {
  public:
    explicit Table(void *node);

    // --- Data Access & Traversal ---
    Cell cell(int row, int col);
    Row row(int rowIndex);
    std::vector<Row> rows() const;
    void mergeCells(int startRow, int startCol, int endRow, int endCol);

    // --- Ergonomic Table Formatting ---
    Table &setBorders(const BorderSettings &all);
    Table &setBorders(const BorderSettings &outer, const BorderSettings &inner);
    Table &setBorders(const BorderSettings &top, const BorderSettings &bottom, const BorderSettings &left,
                      const BorderSettings &right, const BorderSettings &insideH, const BorderSettings &insideV);

    Table &setColumnWidth(int colIndex, int twips);
    Table &setColumnWidths(const std::vector<int> &twipsList);

    Table &setAlignment(gsl::czstring align);

    // --- Data Extraction & Manipulation ---
    int gridSpan() const;
    std::string vMerge() const;
    std::vector<BlockElement> elements() const;
    int replaceText(const std::string &search, const std::string &replace);
    std::string text() const;

    // --- DOM Mutability ---
    void remove();
    Table cloneAfter();
    Paragraph insertParagraphAfter(const std::string &text = "");

  private:
    void *node_;
};

/**
 * @brief Represents a Microsoft Word Document (.docx)
 */

enum class NumberingFormat { Decimal, LowerLetter, UpperLetter, LowerRoman, UpperRoman, Bullet };

struct ListLevel {
    int levelIndex = 0; // 0 to 8
    NumberingFormat format = NumberingFormat::Decimal;
    std::string text;        // e.g. "%1.", "%1.%2.", or bullet char
    std::string jc = "left"; // Justification: left, center, right
    int indentTwips = 720;
    int hangingTwips = 360;
    std::string fontAscii = ""; // Specific font for bullets, e.g. "Symbol"
};

class AbstractNumbering {
  public:
    explicit AbstractNumbering(void *node);
    AbstractNumbering &addLevel(const ListLevel &level);

  private:
    void *node_;
};

class NumberingCollection {
  public:
    explicit NumberingCollection(void *node);

    // --- Abstract Data ---
    AbstractNumbering addAbstractNumbering(int abstractNumId);
    int addList(int abstractNumId, int restartNumId = -1);

    // --- High-Level Ergonomic API ---
    int addBulletList();
    int addNumberedList();

  private:
    void *node_;
};

class TextBox {
  public:
    explicit TextBox(void *node);

    // TextBox acts as a miniature container
    Paragraph addParagraph(const std::string &text = "");
    void addHtml(const std::string &html);

    TextBox &setFillColor(const std::string &hexColor);
    TextBox &setLineColor(const std::string &hexColor);

  private:
    void *node_;
};

enum class ChartType { Bar, Line, Pie, Area, Scatter };

enum class LegendPosition { Bottom, Top, Left, Right, None };

enum class AreaGrouping { Standard, Stacked, PercentStacked };

enum class ScatterStyle { Marker, SmoothMarker, LineMarker };

struct ChartOptions {
    std::string title;
    LegendPosition legendPos = LegendPosition::Bottom;
    bool showDataLabels = false;
    int widthTwips = 8000;  // approx 14 cm
    int heightTwips = 4800; // approx 8 cm

    ScatterStyle scatterStyle = ScatterStyle::Marker;
    bool smoothLines = false;
    AreaGrouping areaGrouping = AreaGrouping::Standard;
};

struct ChartSeries {
    std::string name;
    std::vector<std::string> categories;
    std::vector<double> xValues;
    std::vector<double> values;
    std::string colorHex; // Empty string means auto
};

class Chart {
  public:
    explicit Chart(void *node);
    std::string relId() const;

  private:
    void *node_;
};

/**
 * @brief Represents the main Office Open XML Word document (.docx).
 *
 * This class serves as the main entry point for creating, parsing, and manipulating
 * DOCX files. It manages the core lifecycle, resources, relationships, and global configurations.
 */
class Document {
  public:
    Document();
    explicit Document(const std::string &templatePath);
    ~Document();

    // Disable copy for safety (RAII)
    Document(const Document &) = delete;
    Document &operator=(const Document &) = delete;

    // --- IO Operations ---
    bool save(gsl::czstring filepath);
    bool load(gsl::czstring filepath);
    bool validate(gsl::czstring partName, const SchemaValidator &validator, std::string &outErrors) const;

    void setEvenAndOddHeaders(bool val = true);

    // --- Content Creation ---
    Paragraph addParagraph(const std::string &text = "");
    void addHtml(const std::string &html);

    /**
     * @brief Nests a new table inside this cell.
     * @param rows Initial number of rows.
     * @param cols Initial number of columns.
     * @return The proxy object to the newly created nested table.
     */
    Table addTable(int rows, int cols);

    // --- DOM Traversal & Extraction ---
    std::vector<BlockElement> elements() const;
    Section finalSection();
    std::vector<Paragraph> paragraphs() const;
    std::vector<Table> tables() const;

    std::vector<Chart> charts() const;

    /**
     * @brief Retrieves all footnotes in the document as a map of ID to text.
     */
    std::map<int, std::string> footnotes() const;

    /**
     * @brief Retrieves all endnotes in the document as a map of ID to text.
     */
    std::map<int, std::string> endnotes() const;

    void setMetadata(const Metadata &meta);
    Metadata metadata() const;

    void addTableOfContents(gsl::czstring title = "Table of Contents", int max_levels = 3,
                            TOCLeader leader = TOCLeader::Dot);
    void addWatermark(const std::string &text);

    Chart addChart(ChartType type, const std::vector<ChartSeries> &series, const ChartOptions &options = ChartOptions());

    int createFootnote(const std::string &text);
    /**
     * @brief Creates a new comment in the document's global comments registry.
     * @param text The comment text.
     * @param author The author of the comment (defaults to "Author").
     * @param initials The initials of the author.
     * @return The unique identifier integer for the created comment, used to anchor it to text.
     */
    int createComment(const std::string &text, const std::string &author = "Author", const std::string &initials = "");
    int createEndnote(const std::string &text);

    StyleCollection styles();
    NumberingCollection numbering();

    /**

     * @param styleId Unique ID for the style (e.g., "Heading1")
     * @param name Human-readable name
     */
    void addStyle(gsl::czstring styleId, gsl::czstring name);

    // --- Utilities ---
    std::string convertMathMLToOMML(const std::string &mathml) const;
    std::string convertLaTeXToOMML(const std::string &latex) const;
    int replaceText(const std::string &search, const std::string &replace);
    /**
     * @brief Template Engine: Clones a table row containing the `search` string.
     * Generates a new row for each element in the `values` array, replacing placeholders with the map values.
     * The original row is removed.
     * @param search The placeholder string identifying the target row.
     * @param values A vector of maps for replacements.
     * @return Number of rows generated.
     */
    int cloneRowAndSetValues(const std::string &search, const std::vector<std::map<std::string, std::string>> &values);
    /**
     * @param search The placeholder string identifying the target row.
     */

  private:
    struct Impl;
    std::unique_ptr<Impl> pimpl;
};

} // namespace openword
