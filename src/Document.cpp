#include <openword/Document.h>
#include "Internal.h"

// Standard library
#include <sstream>

// Third-party libraries
#include <pugixml.hpp>
#include <zip.h>
#include <fmt/core.h>

namespace openword {

// --- Document Implementation ---
struct Document::Impl {
    pugi::xml_document doc;
    pugi::xml_node body;
    pugi::xml_document styles_doc;
    pugi::xml_node styles_root;
    RelationshipManager doc_rels;
    bool has_styles = false;

    void initialize() {
        auto decl = doc.append_child(pugi::node_declaration);
        decl.append_attribute("version") = "1.0";
        decl.append_attribute("encoding") = "UTF-8";
        decl.append_attribute("standalone") = "yes";
        auto d = doc.append_child("w:document");
        d.append_attribute("xmlns:w") = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        d.append_attribute("xmlns:r") = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        d.append_attribute("xmlns:wp") = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        d.append_attribute("xmlns:a") = "http://schemas.openxmlformats.org/drawingml/2006/main";
        d.append_attribute("xmlns:pic") = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        body = d.append_child("w:body");
        
        // Add default section properties to make Word happy
        auto sectPr = body.append_child("w:sectPr");
        auto pgSz = sectPr.append_child("w:pgSz");
        pgSz.append_attribute("w:w") = "11906"; // A4
        pgSz.append_attribute("w:h") = "16838";
        auto pgMar = sectPr.append_child("w:pgMar");
        pgMar.append_attribute("w:top") = "1440"; // 1 inch
        pgMar.append_attribute("w:right") = "1440";
        pgMar.append_attribute("w:bottom") = "1440";
        pgMar.append_attribute("w:left") = "1440";
        pgMar.append_attribute("w:header") = "720";
        pgMar.append_attribute("w:footer") = "720";
        pgMar.append_attribute("w:gutter") = "0";

        auto decl2 = styles_doc.append_child(pugi::node_declaration);
        decl2.append_attribute("version") = "1.0";
        decl2.append_attribute("encoding") = "UTF-8";
        decl2.append_attribute("standalone") = "yes";
        styles_root = styles_doc.append_child("w:styles");
        styles_root.append_attribute("xmlns:w") = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    }
};

Document::Document() : pimpl(std::make_unique<Impl>()) { pimpl->initialize(); }
Document::~Document() = default;

Paragraph Document::addParagraph(const std::string& text) {
    auto sectPr = pimpl->body.child("w:sectPr");
    auto p_node = pimpl->body.insert_child_before("w:p", sectPr);
    Paragraph para_proxy(p_node.internal_object());
    if (!text.empty()) {
        auto r_node = p_node.append_child("w:r");
        Run run_proxy(r_node.internal_object());
        run_proxy.setText(text);
    }
    return para_proxy;
}

Table Document::addTable(int rows, int cols) {
    auto sectPr = pimpl->body.child("w:sectPr");
    auto tbl = pimpl->body.insert_child_before("w:tbl", sectPr);
    
    auto tblPr = tbl.append_child("w:tblPr");
    auto tblStyle = tblPr.append_child("w:tblStyle");
    tblStyle.append_attribute("w:val") = "TableGrid";
    auto tblW = tblPr.append_child("w:tblW");
    tblW.append_attribute("w:w") = "0";
    tblW.append_attribute("w:type") = "auto";
    auto tblBorders = tblPr.append_child("w:tblBorders");
    for (auto border : {"w:top", "w:left", "w:bottom", "w:right", "w:insideH", "w:insideV"}) {
        auto b = tblBorders.append_child(border);
        b.append_attribute("w:val") = "single";
        b.append_attribute("w:sz") = "4";
        b.append_attribute("w:space") = "0";
        b.append_attribute("w:color") = "auto";
    }

    auto tblGrid = tbl.append_child("w:tblGrid");
    for (int c = 0; c < cols; ++c) {
        tblGrid.append_child("w:gridCol");
    }

    for (int r = 0; r < rows; ++r) {
        auto tr = tbl.append_child("w:tr");
        for (int c = 0; c < cols; ++c) {
            auto tc = tr.append_child("w:tc");
            auto tcPr = tc.append_child("w:tcPr");
            auto tcW = tcPr.append_child("w:tcW");
            tcW.append_attribute("w:w") = "0";
            tcW.append_attribute("w:type") = "auto";
        }
    }

    return Table(tbl.internal_object());
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
        
        auto media_store = pimpl->doc.child("w:document").child("openword_media");
        int media_count = 0;
        if (media_store) {
            for (auto m : media_store.children("media")) {
                media_count++;
                std::string path = m.attribute("path").value();
                std::string old_rid = m.attribute("rId").value();
                
                std::string ext = "png";
                auto dot_pos = path.find_last_of('.');
                if (dot_pos != std::string::npos) ext = path.substr(dot_pos + 1);
                
                std::string target = "media/image" + std::to_string(media_count) + "." + ext;
                pimpl->doc_rels.addRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target, old_rid);
                
                zip_source_t* img_s = zip_source_file(z, path.c_str(), 0, 0);
                if (img_s) {
                    zip_file_add(z, ("word/" + target).c_str(), img_s, ZIP_FL_OVERWRITE);
                }
            }
            pimpl->doc.child("w:document").remove_child(media_store);
            
            std::stringstream doc_stream2;
            pimpl->doc.save(doc_stream2, "", pugi::format_raw);
            doc_str = doc_stream2.str();
            s = zip_source_buffer(z, doc_str.data(), doc_str.size(), 0);
            zip_file_add(z, "word/document.xml", s, ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
        }

        if (!pimpl->doc_rels.empty()) {
            pugi::xml_document rels_doc;
            auto decl3 = rels_doc.append_child(pugi::node_declaration);
            decl3.append_attribute("version") = "1.0";
            decl3.append_attribute("encoding") = "UTF-8";
            decl3.append_attribute("standalone") = "yes";
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
            "  <Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>\n"
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

bool Document::load(gsl::czstring filepath) {
    try {
        int error = 0;
        zip_t* z = zip_open(filepath, ZIP_RDONLY, &error);
        if (!z) return false;

        pimpl->doc.reset();
        pimpl->styles_doc.reset();
        pimpl->has_styles = false;

        auto read_xml_from_zip = [&](const char* filename, pugi::xml_document& target_doc) -> bool {
            zip_stat_t st;
            zip_stat_init(&st);
            if (zip_stat(z, filename, 0, &st) != 0) return false;

            zip_file_t* f = zip_fopen(z, filename, 0);
            if (!f) return false;

            std::string content(st.size, '\0');
            zip_fread(f, content.data(), st.size);
            zip_fclose(f);

            return target_doc.load_string(content.c_str());
        };

        if (!read_xml_from_zip("word/document.xml", pimpl->doc)) {
            zip_close(z);
            return false;
        }
        pimpl->body = pimpl->doc.child("w:document").child("w:body");

        if (read_xml_from_zip("word/styles.xml", pimpl->styles_doc)) {
            pimpl->has_styles = true;
            pimpl->styles_root = pimpl->styles_doc.child("w:styles");
        }

        zip_close(z);
        return true;
    } catch (...) {
        return false;
    }
}

std::vector<Paragraph> Document::paragraphs() const {
    std::vector<Paragraph> result;
    for (auto node : pimpl->body.children("w:p")) {
        result.push_back(Paragraph(node.internal_object()));
    }
    return result;
}

} // namespace openword
