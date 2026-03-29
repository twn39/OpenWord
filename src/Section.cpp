#include <openword/Document.h>
#include "Internal.h"
#include <string>

namespace openword {

Header::Header(void* node) : node_(node) { Expects(node_ != nullptr); }
Paragraph Header::addParagraph(const std::string& text) {
    auto n = cast_node(node_);
    auto p = n.append_child("w:p");
    Paragraph para(p.internal_object());
    if (!text.empty()) para.addRun(text);
    return para;
}

Footer::Footer(void* node) : node_(node) { Expects(node_ != nullptr); }
Paragraph Footer::addParagraph(const std::string& text) {
    auto n = cast_node(node_);
    auto p = n.append_child("w:p");
    Paragraph para(p.internal_object());
    if (!text.empty()) para.addRun(text);
    return para;
}

Section::Section(void* node) : node_(node) { Expects(node_ != nullptr); }

Section& Section::setPageSize(uint32_t w_twips, uint32_t h_twips, Orientation orient) {
    auto n = cast_node(node_);
    auto pgSz = n.child("w:pgSz");
    if (!pgSz) pgSz = n.append_child("w:pgSz");
    
    auto set_attr = [&](const char* name, uint32_t val) {
        if (!pgSz.attribute(name)) pgSz.append_attribute(name) = std::to_string(val).c_str();
        else pgSz.attribute(name).set_value(std::to_string(val).c_str());
    };
    
    set_attr("w:w", w_twips);
    set_attr("w:h", h_twips);
    
    if (orient == Orientation::Landscape) {
        if (!pgSz.attribute("w:orient")) pgSz.append_attribute("w:orient") = "landscape";
        else pgSz.attribute("w:orient").set_value("landscape");
    } else {
        pgSz.remove_attribute("w:orient");
    }
    return *this;
}

Section& Section::setMargins(const Margins& margins) {
    auto n = cast_node(node_);
    auto pgMar = n.child("w:pgMar");
    if (!pgMar) pgMar = n.append_child("w:pgMar");
    
    auto set_attr = [&](const char* name, uint32_t val) {
        if (!pgMar.attribute(name)) pgMar.append_attribute(name) = std::to_string(val).c_str();
        else pgMar.attribute(name).set_value(std::to_string(val).c_str());
    };
    
    set_attr("w:top", margins.top);
    set_attr("w:right", margins.right);
    set_attr("w:bottom", margins.bottom);
    set_attr("w:left", margins.left);
    set_attr("w:header", margins.header);
    set_attr("w:footer", margins.footer);
    set_attr("w:gutter", margins.gutter);
    return *this;
}

static pugi::xml_node createPart(pugi::xml_node sectPr, const std::string& partType, HeaderFooterType type) {
    auto root = sectPr.root().child("w:document");
    auto parts = root.child("openword_parts");
    if (!parts) parts = root.append_child("openword_parts");
    
    int rel_id = root.attribute("openword_next_rel_id").as_int(100);
    root.attribute("openword_next_rel_id").set_value(std::to_string(rel_id + 1).c_str());
    std::string rid_str = "rId" + std::to_string(rel_id);
    
    std::string filename = partType + std::to_string(rel_id) + ".xml";
    
    auto part = parts.append_child("part");
    part.append_attribute("type") = partType.c_str();
    part.append_attribute("file") = filename.c_str();
    part.append_attribute("rId") = rid_str.c_str();
    
    auto rootNodeName = (partType == "header") ? "w:hdr" : "w:ftr";
    auto node = part.append_child(rootNodeName);
    node.append_attribute("xmlns:w") = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    node.append_attribute("xmlns:r") = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    
    auto refName = (partType == "header") ? "w:headerReference" : "w:footerReference";
    
    // Schema requirement: headerReference and footerReference MUST come BEFORE pgSz, pgMar, etc.
    auto ref = sectPr.prepend_child(refName);
    
    std::string type_str = "default";
    if (type == HeaderFooterType::First) type_str = "first";
    else if (type == HeaderFooterType::Even) type_str = "even";
    
    ref.append_attribute("w:type") = type_str.c_str();
    ref.append_attribute("r:id") = rid_str.c_str();
    
    return node;
}

Header Section::addHeader(HeaderFooterType type) {
    return Header(createPart(cast_node(node_), "header", type).internal_object());
}

Footer Section::addFooter(HeaderFooterType type) {
    return Footer(createPart(cast_node(node_), "footer", type).internal_object());
}

} // namespace openword
