#include <openword/Document.h>

// Standard library
#include <sstream>
#include <map>

// Third-party libraries
#include <pugixml.hpp>
#include <zip.h>
#include <fmt/core.h>

namespace openword {

// Helper to safely cast void* to pugi::xml_node
static pugi::xml_node cast_node(void* ptr) {
    Expects(ptr != nullptr);
    return pugi::xml_node(static_cast<pugi::xml_node_struct*>(ptr));
}

// --- Relationship Manager (Internal) ---
class RelationshipManager {
public:
    std::string addRelationship(const std::string& type, const std::string& target) {
        std::string rId = "rId" + std::to_string(nextId_++);
        relationships_[rId] = {type, target};
        return rId;
    }
    bool empty() const { return relationships_.empty(); }
    void serialize(pugi::xml_node parent) const {
        for (const auto& [id, data] : relationships_) {
            auto rel = parent.append_child("Relationship");
            rel.append_attribute("Id") = id.c_str();
            rel.append_attribute("Type") = data.type.c_str();
            rel.append_attribute("Target") = data.target.c_str();
        }
    }
private:
    struct RelData { std::string type; std::string target; };
    int nextId_ = 1;
    std::map<std::string, RelData> relationships_;
};

// --- Run Implementation ---
Run::Run(void* node) : node_(node) {
    Expects(node_ != nullptr);
}

Run& Run::setText(const std::string& text) {
    auto n = cast_node(node_);
    auto t = n.child("w:t");
    if (!t) t = n.append_child("w:t");
    t.text().set(text.c_str());
    return *this;
}

Run& Run::setBold(bool val) {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr) rPr = n.prepend_child("w:rPr");
    auto b = rPr.child("w:b");
    if (val && !b) rPr.append_child("w:b");
    else if (!val && b) rPr.remove_child(b);
    return *this;
}

Run& Run::setItalic(bool val) {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr) rPr = n.prepend_child("w:rPr");
    auto i = rPr.child("w:i");
    if (val && !i) rPr.append_child("w:i");
    else if (!val && i) rPr.remove_child(i);
    return *this;
}

Run& Run::setFontSize(int halfPoints) {
    auto n = cast_node(node_);
    auto rPr = n.child("w:rPr");
    if (!rPr) rPr = n.prepend_child("w:rPr");
    auto sz = rPr.child("w:sz");
    if (!sz) sz = rPr.append_child("w:sz");
    sz.append_attribute("w:val").set_value(std::to_string(halfPoints).c_str());
    return *this;
}

// --- Paragraph Implementation ---
Paragraph::Paragraph(void* node) : node_(node) {
    Expects(node_ != nullptr);
}

Run Paragraph::addRun(const std::string& text) {
    auto n = cast_node(node_);
    auto r_node = n.append_child("w:r");
    
    Run run_proxy(r_node.internal_object());
    if (!text.empty()) run_proxy.setText(text);
    return run_proxy;
}

Paragraph& Paragraph::setStyle(gsl::czstring styleId) {
    auto n = cast_node(node_);
    auto pPr = n.child("w:pPr");
    if (!pPr) pPr = n.prepend_child("w:pPr");
    auto pStyle = pPr.child("w:pStyle");
    if (!pStyle) pStyle = pPr.prepend_child("w:pStyle");
    pStyle.append_attribute("w:val").set_value(styleId);
    return *this;
}

Paragraph& Paragraph::setAlignment(gsl::czstring align) {
    auto n = cast_node(node_);
    auto pPr = n.child("w:pPr");
    if (!pPr) pPr = n.prepend_child("w:pPr");
    auto jc = pPr.child("w:jc");
    if (!jc) jc = pPr.append_child("w:jc");
    jc.append_attribute("w:val").set_value(align);
    return *this;
}

// --- Document Implementation ---
struct Document::Impl {
    pugi::xml_document doc;
    pugi::xml_node body;
    pugi::xml_document styles_doc;
    pugi::xml_node styles_root;
    RelationshipManager doc_rels;
    bool has_styles = false;

    void initialize() {
        doc.append_child(pugi::node_declaration).append_attribute("version") = "1.0";
        auto d = doc.append_child("w:document");
        d.append_attribute("xmlns:w") = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        d.append_attribute("xmlns:r") = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        body = d.append_child("w:body");

        styles_doc.append_child(pugi::node_declaration).append_attribute("version") = "1.0";
        styles_root = styles_doc.append_child("w:styles");
        styles_root.append_attribute("xmlns:w") = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    }
};

Document::Document() : pimpl(std::make_unique<Impl>()) { pimpl->initialize(); }
Document::~Document() = default;

Paragraph Document::addParagraph(const std::string& text) {
    auto p_node = pimpl->body.append_child("w:p");
    Paragraph para_proxy(p_node.internal_object());
    if (!text.empty()) {
        auto r_node = p_node.append_child("w:r");
        Run run_proxy(r_node.internal_object());
        run_proxy.setText(text);
    }
    return para_proxy;
}

void Document::addStyle(gsl::czstring styleId, gsl::czstring name) {
    pimpl->has_styles = true;
    auto style = pimpl->styles_root.append_child("w:style");
    style.append_attribute("w:type") = "paragraph";
    style.append_attribute("w:styleId") = styleId;
    auto nameNode = style.append_child("w:name");
    nameNode.append_attribute("w:val") = name;
    
    auto rPr = style.append_child("w:rPr");
    rPr.append_child("w:b");
    auto sz = rPr.append_child("w:sz");
    sz.append_attribute("w:val") = "48"; // 24pt
    auto color = rPr.append_child("w:color");
    color.append_attribute("w:val") = "FF0000"; // Red
}

bool Document::save(gsl::czstring filepath) {
    try {
        int error = 0;
        zip_t* z = zip_open(filepath, ZIP_CREATE | ZIP_TRUNCATE, &error);
        if (!z) return false;
        
        std::stringstream doc_stream;
        pimpl->doc.save(doc_stream, "", pugi::format_raw);
        std::string doc_str = doc_stream.str();
        
        std::string styles_str;
        std::string rels_str;
        
        zip_source_t* s = zip_source_buffer(z, doc_str.data(), doc_str.size(), 0);
        zip_file_add(z, "word/document.xml", s, ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
        
        if (pimpl->has_styles) {
            std::stringstream styles_stream;
            pimpl->styles_doc.save(styles_stream, "", pugi::format_raw);
            styles_str = styles_stream.str();
            
            zip_source_t* ss = zip_source_buffer(z, styles_str.data(), styles_str.size(), 0);
            zip_file_add(z, "word/styles.xml", ss, ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
            
            pimpl->doc_rels.addRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                "styles.xml"
            );
        }
        
        if (!pimpl->doc_rels.empty()) {
            pugi::xml_document rels_doc;
            rels_doc.append_child(pugi::node_declaration).append_attribute("version") = "1.0";
            auto rels_root = rels_doc.append_child("Relationships");
            rels_root.append_attribute("xmlns") = "http://schemas.openxmlformats.org/package/2006/relationships";
            
            pimpl->doc_rels.serialize(rels_root);
            
            std::stringstream rels_stream;
            rels_doc.save(rels_stream, "", pugi::format_raw);
            rels_str = rels_stream.str();
            
            zip_source_t* rs = zip_source_buffer(z, rels_str.data(), rels_str.size(), 0);
            zip_file_add(z, "word/_rels/document.xml.rels", rs, ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
        }
        
        // Generate [Content_Types].xml
        std::string content_types = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
            "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
            "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
            "  <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\n";
        if (pimpl->has_styles) {
            content_types += "  <Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>\n";
        }
        content_types += "</Types>";
        
        zip_source_t* ct_s = zip_source_buffer(z, content_types.data(), content_types.size(), 0);
        zip_file_add(z, "[Content_Types].xml", ct_s, ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);

        // Generate _rels/.rels
        std::string root_rels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
            "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>\n"
            "</Relationships>";
            
        zip_source_t* rr_s = zip_source_buffer(z, root_rels.data(), root_rels.size(), 0);
        zip_file_add(z, "_rels/.rels", rr_s, ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);

        zip_close(z);
        return true;
    } catch(...) {
        return false;
    }
}

std::string Document::convertMathMLToOMML(const std::string& mathml) const {
    // Mock representation
    if(mathml.find("<math") != std::string::npos) {
        return "<m:oMath></m:oMath>";
    }
    return "";
}

} // namespace openword
