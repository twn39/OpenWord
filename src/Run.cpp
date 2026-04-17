#include "Internal.h"
#include <openword/Document.h>

#include <fmt/core.h>
#include <pugixml.hpp>

namespace openword {

std::string Color::hex() const {
    if (is_auto)
        return "auto";
    return fmt::format("{:02X}{:02X}{:02X}", r, g, b);
}

Run::Run(void *node) : node_(node) { Expects(node_ != nullptr); }

Run &Run::setText(const std::string &text) {
    auto n = cast_node(node_);
    auto t = n.child("w:t");
    if (!t)
        t = n.append_child("w:t");
    t.text().set(text.c_str());
    if (!text.empty() && (text.front() == ' ' || text.back() == ' ' || text.find('\t') != std::string::npos ||
                          text.find('\n') != std::string::npos)) {
        if (!t.attribute("xml:space")) {
            t.append_attribute("xml:space") = "preserve";
        }
    }
    return *this;
}

Run &Run::setBold(bool val) {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        rPr = n.prepend_child("w:rPr");
    auto b = rPr.child("w:b");
    if (val && !b)
        rPr.append_child("w:b");
    else if (!val && b)
        rPr.remove_child(b);
    return *this;
}

Run &Run::setItalic(bool val) {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        rPr = n.prepend_child("w:rPr");
    auto i = rPr.child("w:i");
    if (val && !i)
        rPr.append_child("w:i");
    else if (!val && i)
        rPr.remove_child(i);
    return *this;
}

Run &Run::setFontSize(int halfPoints) {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        rPr = n.prepend_child("w:rPr");
    auto sz = rPr.child("w:sz");
    if (!sz)
        sz = rPr.append_child("w:sz");
    sz.append_attribute("w:val").set_value(std::to_string(halfPoints).c_str());
    return *this;
}

Run &Run::setHighlight(gsl::czstring color) {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        rPr = n.prepend_child("w:rPr");
    auto highlight = rPr.child("w:highlight");
    if (!highlight)
        highlight = rPr.append_child("w:highlight");
    highlight.append_attribute("w:val").set_value(color);
    return *this;
}

Run &Run::setShading(const Color &fillColor) {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        rPr = n.prepend_child("w:rPr");
    auto shd = rPr.child("w:shd");
    if (!shd)
        shd = rPr.append_child("w:shd");
    shd.append_attribute("w:val").set_value("clear");
    shd.append_attribute("w:fill").set_value(fillColor.hex().c_str());
    return *this;
}

Run &Run::setColor(const Color &color) {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        rPr = n.prepend_child("w:rPr");
    auto c = rPr.child("w:color");
    if (!c)
        c = rPr.append_child("w:color");
    c.append_attribute("w:val").set_value(color.hex().c_str());
    return *this;
}

Run &Run::setFontFamily(gsl::czstring ascii, gsl::czstring eastAsia) {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        rPr = n.prepend_child("w:rPr");
    auto fonts = rPr.child("w:rFonts");
    if (!fonts)
        fonts = rPr.append_child("w:rFonts");
    fonts.append_attribute("w:ascii").set_value(ascii);
    fonts.append_attribute("w:hAnsi").set_value(ascii);
    if (eastAsia && eastAsia[0] != '\0') {
        fonts.append_attribute("w:eastAsia").set_value(eastAsia);
    }
    return *this;
}

Run &Run::setStrike(bool val) {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        rPr = n.prepend_child("w:rPr");
    auto strike = rPr.child("w:strike");
    if (val && !strike)
        rPr.append_child("w:strike");
    else if (!val && strike)
        rPr.remove_child(strike);
    return *this;
}

Run &Run::setDoubleStrike(bool val) {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        rPr = n.prepend_child("w:rPr");
    auto dstrike = rPr.child("w:dstrike");
    if (val && !dstrike)
        rPr.append_child("w:dstrike");
    else if (!val && dstrike)
        rPr.remove_child(dstrike);
    return *this;
}

Run &Run::setVertAlign(VertAlign align) {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        rPr = n.prepend_child("w:rPr");
    auto vertAlign = rPr.child("w:vertAlign");

    if (align == VertAlign::Baseline) {
        if (vertAlign)
            rPr.remove_child(vertAlign);
    } else {
        if (!vertAlign)
            vertAlign = rPr.append_child("w:vertAlign");
        vertAlign.append_attribute("w:val").set_value(align == VertAlign::Superscript ? "superscript" : "subscript");
    }
    return *this;
}

Run &Run::setUnderline(gsl::czstring val) {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        rPr = n.prepend_child("w:rPr");
    auto u = rPr.child("w:u");
    if (!u)
        u = rPr.append_child("w:u");
    u.append_attribute("w:val").set_value(val);
    return *this;
}

std::string Run::text() const {
    auto n = cast_node(node_);
    auto t = n.child("w:t");
    return t ? t.text().get() : "";
}

static const char *highlightColorToString(HighlightColor color) {
    switch (color) {
    case HighlightColor::Yellow:
        return "yellow";
    case HighlightColor::Green:
        return "green";
    case HighlightColor::Cyan:
        return "cyan";
    case HighlightColor::Magenta:
        return "magenta";
    case HighlightColor::Blue:
        return "blue";
    case HighlightColor::Red:
        return "red";
    case HighlightColor::DarkBlue:
        return "darkBlue";
    case HighlightColor::DarkCyan:
        return "darkCyan";
    case HighlightColor::DarkGreen:
        return "darkGreen";
    case HighlightColor::DarkMagenta:
        return "darkMagenta";
    case HighlightColor::DarkRed:
        return "darkRed";
    case HighlightColor::DarkYellow:
        return "darkYellow";
    case HighlightColor::DarkGray:
        return "darkGray";
    case HighlightColor::LightGray:
        return "lightGray";
    case HighlightColor::Black:
        return "black";
    default:
        return "none";
    }
}

Run &Run::setHighlight(HighlightColor color) {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        rPr = n.prepend_child("w:rPr");

    auto highlight = rPr.child("w:highlight");
    if (!highlight)
        highlight = rPr.append_child("w:highlight");

    highlight.remove_attribute("w:val");
    if (color != HighlightColor::Default) {
        highlight.append_attribute("w:val") = highlightColorToString(color);
    } else {
        rPr.remove_child(highlight);
    }

    return *this;
}

Run &Run::setCharacterSpacing(int twips) {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        rPr = n.prepend_child("w:rPr");

    auto spacing = rPr.child("w:spacing");
    if (!spacing)
        spacing = rPr.append_child("w:spacing");

    spacing.remove_attribute("w:val");
    if (twips != 0) {
        spacing.append_attribute("w:val") = std::to_string(twips).c_str();
    } else {
        rPr.remove_child(spacing);
    }

    return *this;
}

bool Run::isBold() const {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        return false;
    return rPr.child("w:b") || rPr.child("w:bCs");
}

bool Run::isItalic() const {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        return false;
    return rPr.child("w:i") || rPr.child("w:iCs");
}

bool Run::isStrike() const {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        return false;
    auto strike = rPr.child("w:strike");
    if (strike) {
        auto val = strike.attribute("w:val");
        return val.empty() || std::string(val.value()) == "true" || std::string(val.value()) == "1";
    }
    return false;
}

HighlightColor Run::highlightColor() const {
    auto n = cast_node(node_);
    auto hl = n.child("w:rPr").child("w:highlight");
    if (hl) {
        std::string v = hl.attribute("w:val").value();
        if (v == "yellow")
            return HighlightColor::Yellow;
        if (v == "green")
            return HighlightColor::Green;
        if (v == "cyan")
            return HighlightColor::Cyan;
        if (v == "magenta")
            return HighlightColor::Magenta;
        if (v == "blue")
            return HighlightColor::Blue;
        if (v == "red")
            return HighlightColor::Red;
        if (v == "darkBlue")
            return HighlightColor::DarkBlue;
        if (v == "darkCyan")
            return HighlightColor::DarkCyan;
        if (v == "darkGreen")
            return HighlightColor::DarkGreen;
        if (v == "darkMagenta")
            return HighlightColor::DarkMagenta;
        if (v == "darkRed")
            return HighlightColor::DarkRed;
        if (v == "darkYellow")
            return HighlightColor::DarkYellow;
        if (v == "darkGray")
            return HighlightColor::DarkGray;
        if (v == "lightGray")
            return HighlightColor::LightGray;
        if (v == "black")
            return HighlightColor::Black;
    }
    return HighlightColor::Default;
}

} // namespace openword

openword::Run &openword::Run::addComment(int commentId) {
    auto n = cast_node(node_);
    auto parent = n.parent();

    std::string id_str = std::to_string(commentId);

    // Insert w:commentRangeStart BEFORE this Run
    auto startNode = parent.insert_child_before("w:commentRangeStart", n);
    startNode.append_attribute("w:id") = id_str.c_str();

    // Insert w:commentRangeEnd AFTER this Run
    auto endNode = parent.insert_child_after("w:commentRangeEnd", n);
    endNode.append_attribute("w:id") = id_str.c_str();

    // Insert the reference tag w:commentReference immediately after the RangeEnd
    auto refRun = parent.insert_child_after("w:r", endNode);
    refRun.append_child("w:rPr").append_child("w:rStyle").append_attribute("w:val") = "CommentReference";
    refRun.append_child("w:commentReference").append_attribute("w:id") = id_str.c_str();

    return *this;
}

openword::Run &openword::Run::addLineBreak() {
    auto n = cast_node(node_);
    n.append_child("w:br");
    return *this;
}

openword::Run &openword::Run::addPageBreak() {
    auto n = cast_node(node_);
    auto br = n.append_child("w:br");
    br.append_attribute("w:type") = "page";
    return *this;
}

bool openword::Run::isFootnoteReference() const {
    auto n = cast_node(node_);
    return static_cast<bool>(n.child("w:footnoteReference"));
}

int openword::Run::footnoteId() const {
    auto n = cast_node(node_);
    auto ref = n.child("w:footnoteReference");
    if (ref) {
        return ref.attribute("w:id").as_int(-1);
    }
    return -1;
}

bool openword::Run::isEndnoteReference() const {
    auto n = cast_node(node_);
    return static_cast<bool>(n.child("w:endnoteReference"));
}

int openword::Run::endnoteId() const {
    auto n = cast_node(node_);
    auto ref = n.child("w:endnoteReference");
    if (ref) {
        return ref.attribute("w:id").as_int(-1);
    }
    return -1;
}
