#include "openword/MarkdownConverter.h"
#include "openword/Document.h"

#include <algorithm>
#include <string>
#include <type_traits>
#include <variant>
#include <vector>

namespace openword {

namespace {

std::string renderParagraph(const Paragraph &para, bool inTable = false) {
    std::string out;

    if (!inTable) {
        std::string style = para.styleId();
        if (style.size() == 8 && style.find("Heading") == 0) {
            int headingLevel = style[7] - '0';
            if (headingLevel >= 1 && headingLevel <= 6) {
                out += std::string(headingLevel, '#') + " ";
            }
        } else if (para.isList()) {
            int level = std::max(0, para.listLevel());
            out += std::string(level * 2, ' ') + "- ";
        }
    }

    auto runs = para.runs();
    for (const auto &run : runs) {
        if (run.isFootnoteReference()) {
            int fnId = run.footnoteId();
            if (fnId >= 0) {
                out += "[^" + std::to_string(fnId) + "]";
            }
            continue;
        }

        if (run.isEndnoteReference()) {
            int enId = run.endnoteId();
            if (enId >= 0) {
                out += "[^e_" + std::to_string(enId) + "]";
            }
            continue;
        }

        std::string runText = run.text();
        if (runText.empty())
            continue;

        bool bold = run.isBold();
        bool italic = run.isItalic();
        bool strike = run.isStrike();

        std::string prefix, suffix;
        if (bold) {
            prefix += "**";
            suffix = "**" + suffix;
        }
        if (italic) {
            prefix += "*";
            suffix = "*" + suffix;
        }
        if (strike) {
            prefix += "~~";
            suffix = "~~" + suffix;
        }

        out += prefix + runText + suffix;
    }
    return out;
}

std::string renderBlockElement(const BlockElement &el, bool inTable = false);

std::string renderTable(const Table &table) {
    std::string out;
    auto rows = table.rows();
    if (rows.empty())
        return out;

    for (size_t i = 0; i < rows.size(); ++i) {
        auto cells = rows[i].cells();
        out += "|";
        for (const auto &cell : cells) {
            std::string cellText;
            auto elements = cell.elements();
            for (size_t k = 0; k < elements.size(); ++k) {
                std::string elText = renderBlockElement(elements[k], true);
                if (!elText.empty()) {
                    if (!cellText.empty())
                        cellText += "<br>";
                    cellText += elText;
                }
            }
            // Clean up any remaining newlines that break markdown table formatting
            std::replace(cellText.begin(), cellText.end(), '\n', ' ');
            out += " " + cellText + " |";
        }
        out += "\n";

        if (i == 0) {
            out += "|";
            for (size_t j = 0; j < cells.size(); ++j) {
                out += "---|";
            }
            out += "\n";
        }
    }
    return out;
}

std::string renderBlockElement(const BlockElement &el, bool inTable) {
    return std::visit(
        [inTable](const auto &item) -> std::string {
            using T = std::decay_t<decltype(item)>;
            if constexpr (std::is_same_v<T, Paragraph>) {
                return renderParagraph(item, inTable);
            } else if constexpr (std::is_same_v<T, Table>) {
                if (inTable) {
                    return "[Nested Table unsupported]";
                }
                return renderTable(item);
            } else if constexpr (std::is_same_v<T, Section>) {
                return inTable ? "" : "---";
            }
            return "";
        },
        el);
}

} // namespace

std::string convertToMarkdown(const Document &doc) {
    std::string out;
    auto elements = doc.elements();
    for (const auto &el : elements) {
        std::string rendered = renderBlockElement(el, false);
        if (!rendered.empty()) {
            out += rendered + "\n\n";
        }
    }

    // Append footnotes
    auto footnotes = doc.footnotes();
    if (!footnotes.empty()) {
        out += "---\n\n";
        for (const auto &[id, text] : footnotes) {
            out += "[^" + std::to_string(id) + "]: " + text + "\n";
        }
        out += "\n";
    }

    // Append endnotes
    auto endnotes = doc.endnotes();
    if (!endnotes.empty()) {
        out += "---\n\n";
        for (const auto &[id, text] : endnotes) {
            out += "[^e_" + std::to_string(id) + "]: " + text + "\n";
        }
        out += "\n";
    }

    // Trim trailing newlines
    while (out.size() >= 2 && out.substr(out.size() - 2) == "\n\n") {
        out.pop_back();
    }
    if (!out.empty() && out.back() != '\n')
        out += '\n';
    return out;
}

} // namespace openword