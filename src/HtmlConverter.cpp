#include "openword/Document.h"
#include <optional>

#include <iostream>

#include "Internal.h"

#include <libxml/HTMLparser.h>
#include <libxml/HTMLtree.h>
#include <algorithm>
#include <cctype>
#include <functional>
#include <string>
#include <regex>
#include <sstream>

namespace openword {

namespace {

struct Formatting {
    bool bold = false;
    bool italic = false;
    bool underline = false;
    bool strike = false;
    bool sup = false;
    bool sub = false;
    bool isHeading = false;
    std::string headingStyle;
    std::string linkUrl;
    Color color = Color::Auto();
};

std::string trimAndCollapseWhitespace(const std::string& input) {
    std::string output;
    bool inWhitespace = false;
    for (char c : input) {
        if (std::isspace(static_cast<unsigned char>(c))) {
            if (!inWhitespace) {
                output += ' ';
                inWhitespace = true;
            }
        } else {
            output += c;
            inWhitespace = false;
        }
    }
    // Trim leading/trailing spaces
    return output;
}

Color parseCssColor(const std::string& styleStr) {
    std::regex colorRegex(R"(color\s*:\s*#([0-9a-fA-F]{6}))");
    std::smatch match;
    if (std::regex_search(styleStr, match, colorRegex)) {
        std::string hex = match[1].str();
        uint8_t r = std::stoi(hex.substr(0, 2), nullptr, 16);
        uint8_t g = std::stoi(hex.substr(2, 2), nullptr, 16);
        uint8_t b = std::stoi(hex.substr(4, 2), nullptr, 16);
        return Color(r, g, b);
    }
    return Color::Auto();
}

void traverseHtmlNode(xmlNodePtr node, Formatting fmt, std::optional<Paragraph>& currentPara, bool& hasPara, const std::function<Paragraph()>& createPara) {
    for (xmlNodePtr cur = node->children; cur; cur = cur->next) {
        if (cur->type == XML_TEXT_NODE) {
            std::string text = reinterpret_cast<const char*>(cur->content);
            text = trimAndCollapseWhitespace(text);
            
            if (text.empty() || (text == " " && !hasPara)) continue;

            if (!hasPara) {
                currentPara = createPara();
                if (fmt.isHeading && !fmt.headingStyle.empty()) {
                    currentPara->setStyle(fmt.headingStyle.c_str());
                }
                hasPara = true;
            }

            Run r = fmt.linkUrl.empty() ? currentPara->addRun(text) : currentPara->addHyperlink(text.c_str(), fmt.linkUrl.c_str());
            if (fmt.bold) r.setBold(true);
            if (fmt.italic) r.setItalic(true);
            if (fmt.underline) r.setUnderline("single");
            if (fmt.strike) r.setStrike(true);
            if (fmt.sup) r.setVertAlign(VertAlign::Superscript);
            if (fmt.sub) r.setVertAlign(VertAlign::Subscript);
            if (!fmt.color.is_auto) r.setColor(fmt.color);

        } else if (cur->type == XML_ELEMENT_NODE) {
            std::string tag = reinterpret_cast<const char*>(cur->name);
            std::transform(tag.begin(), tag.end(), tag.begin(), ::tolower);

            Formatting newFmt = fmt;
            bool isBlock = false;

            if (tag == "b" || tag == "strong") newFmt.bold = true;
            else if (tag == "i" || tag == "em") newFmt.italic = true;
            else if (tag == "u") newFmt.underline = true;
            else if (tag == "s" || tag == "strike" || tag == "del") newFmt.strike = true;
            else if (tag == "sup") newFmt.sup = true;
            else if (tag == "sub") newFmt.sub = true;
            else if (tag == "a") {
                xmlChar* href = xmlGetProp(cur, BAD_CAST "href");
                if (href) {
                    newFmt.linkUrl = reinterpret_cast<const char*>(href);
                    xmlFree(href);
                }
            } else if (tag == "span") {
                xmlChar* styleStr = xmlGetProp(cur, BAD_CAST "style");
                if (styleStr) {
                    Color c = parseCssColor(reinterpret_cast<const char*>(styleStr));
                    if (!c.is_auto) newFmt.color = c;
                    xmlFree(styleStr);
                }
            } else if (tag == "br") {
                if (!hasPara) {
                    currentPara = createPara();
                    hasPara = true;
                }
                currentPara->addRun("").addLineBreak();
                continue;
            } else if (tag == "p" || tag == "div" || tag == "h1" || tag == "h2" || tag == "h3" || tag == "h4" || tag == "h5" || tag == "h6" || tag == "li") {
                isBlock = true;
                hasPara = false; // Force new paragraph for the block
                if (tag == "h1") { newFmt.isHeading = true; newFmt.headingStyle = "Heading1"; }
                else if (tag == "h2") { newFmt.isHeading = true; newFmt.headingStyle = "Heading2"; }
                else if (tag == "h3") { newFmt.isHeading = true; newFmt.headingStyle = "Heading3"; }
                else if (tag == "h4") { newFmt.isHeading = true; newFmt.headingStyle = "Heading4"; }
                else if (tag == "h5") { newFmt.isHeading = true; newFmt.headingStyle = "Heading5"; }
                else if (tag == "h6") { newFmt.isHeading = true; newFmt.headingStyle = "Heading6"; }
            }

            traverseHtmlNode(cur, newFmt, currentPara, hasPara, createPara);

            if (isBlock) {
                hasPara = false; // End of block means next text gets a new paragraph
            }
        }
    }
}

} // namespace

void parseHtmlAndInsert(const std::string& html, const std::function<Paragraph()>& createPara) {
    if (html.empty()) return;

    htmlDocPtr doc = htmlReadMemory(html.c_str(), static_cast<int>(html.length()), nullptr, "UTF-8", HTML_PARSE_NOERROR | HTML_PARSE_NOWARNING | HTML_PARSE_RECOVER);
    if (!doc) return;

    xmlNodePtr root = xmlDocGetRootElement(doc);
    if (root) {
        std::optional<Paragraph> currentPara = std::nullopt;
        bool hasPara = false;
        Formatting defaultFmt;
        traverseHtmlNode(root, defaultFmt, currentPara, hasPara, createPara);
    }

    xmlFreeDoc(doc);
}

} // namespace openword
