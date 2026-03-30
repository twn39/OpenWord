#include "Internal.h"
#include "openword/Document.h"
#include <string>

namespace openword {

Font::Font(void *node) : node_(node) {}

Font &Font::setSize(int halfPoints) {
    auto n = cast_node(node_);
    auto sz = n.child("w:sz");
    if (!sz)
        sz = n.append_child("w:sz");
    sz.remove_attribute("w:val");
    sz.append_attribute("w:val") = std::to_string(halfPoints).c_str();

    auto szCs = n.child("w:szCs");
    if (!szCs)
        szCs = n.append_child("w:szCs");
    szCs.remove_attribute("w:val");
    szCs.append_attribute("w:val") = std::to_string(halfPoints).c_str();
    return *this;
}

Font &Font::setBold(bool val) {
    auto n = cast_node(node_);
    auto b = n.child("w:b");
    if (val) {
        if (!b)
            n.append_child("w:b");
        auto bCs = n.child("w:bCs");
        if (!bCs)
            n.append_child("w:bCs");
    } else {
        n.remove_child("w:b");
        n.remove_child("w:bCs");
    }
    return *this;
}

Font &Font::setItalic(bool val) {
    auto n = cast_node(node_);
    auto i = n.child("w:i");
    if (val) {
        if (!i)
            n.append_child("w:i");
        auto iCs = n.child("w:iCs");
        if (!iCs)
            n.append_child("w:iCs");
    } else {
        n.remove_child("w:i");
        n.remove_child("w:iCs");
    }
    return *this;
}

Font &Font::setColor(const std::string &hexColor) {
    auto n = cast_node(node_);
    auto c = n.child("w:color");
    if (!c)
        c = n.append_child("w:color");
    c.remove_attribute("w:val");
    c.append_attribute("w:val") = hexColor.c_str();
    return *this;
}

Font &Font::setName(const std::string &ascii) {
    auto n = cast_node(node_);
    auto rFonts = n.child("w:rFonts");
    if (!rFonts)
        rFonts = n.append_child("w:rFonts");
    rFonts.remove_attribute("w:ascii");
    rFonts.append_attribute("w:ascii") = ascii.c_str();
    rFonts.remove_attribute("w:hAnsi");
    rFonts.append_attribute("w:hAnsi") = ascii.c_str();
    return *this;
}

ParagraphFormat::ParagraphFormat(void *node) : node_(node) {}

ParagraphFormat &ParagraphFormat::setOutlineLevel(int level) {
    auto n = cast_node(node_);
    auto outline = n.child("w:outlineLvl");
    if (!outline)
        outline = n.append_child("w:outlineLvl");
    outline.remove_attribute("w:val");
    outline.append_attribute("w:val") = std::to_string(level).c_str();
    return *this;
}

ParagraphFormat &ParagraphFormat::setSpacing(int beforeTwips, int afterTwips) {
    auto n = cast_node(node_);
    auto spacing = n.child("w:spacing");
    if (!spacing)
        spacing = n.append_child("w:spacing");
    spacing.remove_attribute("w:before");
    spacing.append_attribute("w:before") = std::to_string(beforeTwips).c_str();
    spacing.remove_attribute("w:after");
    spacing.append_attribute("w:after") = std::to_string(afterTwips).c_str();
    return *this;
}

static pugi::xml_node get_or_insert_meta(pugi::xml_node parent, const char *name) {
    if (auto n = parent.child(name)) {
        return n;
    }

    // Find the first layout property node to insert BEFORE it
    pugi::xml_node ref;
    for (auto child : parent.children()) {
        std::string n = child.name();
        if (n == "w:pPr" || n == "w:rPr" || n == "w:tblPr" || n == "w:tcPr" || n == "w:trPr") {
            ref = child;
            break;
        }
    }

    if (ref) {
        return parent.insert_child_before(name, ref);
    }
    return parent.append_child(name);
}

Style::Style(void *node) : node_(node) {}

Font Style::getFont() {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr)
        rPr = n.append_child("w:rPr");
    return Font(rPr.internal_object());
}

ParagraphFormat Style::getParagraphFormat() {
    auto n = cast_node(node_);
    auto pPr = n.child("w:pPr");
    if (!pPr)
        pPr = n.append_child("w:pPr");
    return ParagraphFormat(pPr.internal_object());
}

Style &Style::setName(const std::string &name) {
    auto n = cast_node(node_);
    auto nameNode = get_or_insert_meta(n, "w:name");
    nameNode.remove_attribute("w:val");
    nameNode.append_attribute("w:val") = name.c_str();
    return *this;
}

Style &Style::setBasedOn(const std::string &parentStyleId) {
    auto n = cast_node(node_);
    auto node = get_or_insert_meta(n, "w:basedOn");
    node.remove_attribute("w:val");
    node.append_attribute("w:val") = parentStyleId.c_str();
    return *this;
}

Style &Style::setNextStyle(const std::string &nextStyleId) {
    auto n = cast_node(node_);
    auto node = get_or_insert_meta(n, "w:next");
    node.remove_attribute("w:val");
    node.append_attribute("w:val") = nextStyleId.c_str();
    return *this;
}

Style &Style::setPrimary(bool isPrimary) {
    auto n = cast_node(node_);
    if (isPrimary) {
        get_or_insert_meta(n, "w:qFormat");
    } else {
        n.remove_child("w:qFormat");
    }
    return *this;
}

Style &Style::setUiPriority(int priority) {
    auto n = cast_node(node_);
    auto node = get_or_insert_meta(n, "w:uiPriority");
    node.remove_attribute("w:val");
    node.append_attribute("w:val") = std::to_string(priority).c_str();
    return *this;
}

Style &Style::setHidden(bool isHidden) {
    auto n = cast_node(node_);
    if (isHidden) {
        get_or_insert_meta(n, "w:hidden");
        get_or_insert_meta(n, "w:semiHidden");
    } else {
        n.remove_child("w:hidden");
        n.remove_child("w:semiHidden");
    }
    return *this;
}

StyleCollection::StyleCollection(void *node) : node_(node) {}

Font StyleCollection::getDefaultFont() {
    auto n = cast_node(node_);
    auto docDefaults = n.child("w:docDefaults");
    if (!docDefaults)
        docDefaults = n.prepend_child("w:docDefaults");

    auto rPrDefault = docDefaults.child("w:rPrDefault");
    if (!rPrDefault)
        rPrDefault = docDefaults.append_child("w:rPrDefault");

    auto rPr = rPrDefault.child("w:rPr");
    if (!rPr)
        rPr = rPrDefault.append_child("w:rPr");

    return Font(rPr.internal_object());
}

ParagraphFormat StyleCollection::getDefaultParagraphFormat() {
    auto n = cast_node(node_);
    auto docDefaults = n.child("w:docDefaults");
    if (!docDefaults)
        docDefaults = n.prepend_child("w:docDefaults");

    auto pPrDefault = docDefaults.child("w:pPrDefault");
    if (!pPrDefault)
        pPrDefault = docDefaults.append_child("w:pPrDefault");

    auto pPr = pPrDefault.child("w:pPr");
    if (!pPr)
        pPr = pPrDefault.append_child("w:pPr");

    return ParagraphFormat(pPr.internal_object());
}

Style StyleCollection::get(const std::string &styleId) {
    auto n = cast_node(node_);
    for (auto style : n.children("w:style")) {
        if (std::string(style.attribute("w:styleId").value()) == styleId) {
            return Style(style.internal_object());
        }
    }
    return add(styleId);
}

Style StyleCollection::add(const std::string &styleId, const std::string &type) {
    auto n = cast_node(node_);
    auto style = n.append_child("w:style");
    style.append_attribute("w:type") = type.c_str();
    style.append_attribute("w:styleId") = styleId.c_str();
    return Style(style.internal_object());
}

} // namespace openword
