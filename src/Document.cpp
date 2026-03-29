#include <list>
#include <set>
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
        d.append_attribute("xmlns:wpc") = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
        d.append_attribute("xmlns:mc") = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        d.append_attribute("xmlns:o") = "urn:schemas-microsoft-com:office:office";
        d.append_attribute("xmlns:r") = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        d.append_attribute("xmlns:m") = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        d.append_attribute("xmlns:v") = "urn:schemas-microsoft-com:vml";
        d.append_attribute("xmlns:wp14") = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
        d.append_attribute("xmlns:wp") = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        d.append_attribute("xmlns:w10") = "urn:schemas-microsoft-com:office:word";
        d.append_attribute("xmlns:w") = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        d.append_attribute("xmlns:w14") = "http://schemas.microsoft.com/office/word/2010/wordml";
        d.append_attribute("xmlns:wpg") = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
        d.append_attribute("xmlns:wpi") = "http://schemas.microsoft.com/office/word/2010/wordprocessingInk";
        d.append_attribute("xmlns:wne") = "http://schemas.microsoft.com/office/word/2006/wordml";
        d.append_attribute("xmlns:wps") = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
        d.append_attribute("xmlns:pic") = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        d.append_attribute("xmlns:a") = "http://schemas.openxmlformats.org/drawingml/2006/main";
        d.append_attribute("mc:Ignorable") = "w14 wp14";
        d.append_attribute("openword_next_rel_id") = "100";
        body = d.append_child("w:body");
        
        auto sectPr = body.append_child("w:sectPr");
        auto pgSz = sectPr.append_child("w:pgSz");
        pgSz.append_attribute("w:w") = "11906";
        pgSz.append_attribute("w:h") = "16838";
        auto pgMar = sectPr.append_child("w:pgMar");
        pgMar.append_attribute("w:top") = "1440";
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
    for (int c = 0; c < cols; ++c) tblGrid.append_child("w:gridCol");
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
    sz.append_attribute("w:val") = "48";
    auto color = rPr.append_child("w:color");
    color.append_attribute("w:val") = "FF0000";
}

bool Document::save(gsl::czstring filepath) {
    try {
        int error = 0;
        zip_t* z = zip_open(filepath, ZIP_CREATE | ZIP_TRUNCATE, &error);
        if (!z) return false;
        
        std::list<std::string> zip_buffers; 
        int start_id = pimpl->doc.child("w:document").attribute("openword_next_rel_id").as_int(100);
        pimpl->doc_rels.setNextId(start_id);

        std::string ct = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                         "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
                         "  <Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>\n"
                         "  <Default Extension=\"png\" ContentType=\"image/png\"/>\n"
                         "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
                         "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
                         "  <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\n";

        if (pimpl->has_styles) {
            std::stringstream ss;
            pimpl->styles_doc.save(ss, "", pugi::format_raw);
            zip_buffers.push_back(ss.str());
            zip_file_add(z, "word/styles.xml", zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0), ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
            pimpl->doc_rels.addRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml");
            ct += "  <Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>\n";
        }

        auto media_store = pimpl->doc.child("w:document").child("openword_media");
        if (media_store) {
            int media_count = 0;
            for (auto m : media_store.children("media")) {
                media_count++;
                std::string path = m.attribute("path").value();
                std::string old_rid = m.attribute("rId").value();
                std::string ext = "png";
                auto dot_pos = path.find_last_of('.');
                if (dot_pos != std::string::npos) ext = path.substr(dot_pos + 1);
                std::string target = "media/image" + std::to_string(media_count) + "." + ext;
                pimpl->doc_rels.addRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target, old_rid);
                zip_file_add(z, ("word/" + target).c_str(), zip_source_file(z, path.c_str(), 0, 0), ZIP_FL_OVERWRITE);
            }
            pimpl->doc.child("w:document").remove_child(media_store);
        }

        auto parts = pimpl->doc.child("w:document").child("openword_parts");
        if (parts) {
            for (auto p : parts.children("part")) {
                std::string ptype = p.attribute("type").value();
                std::string pfile = p.attribute("file").value();
                std::string prid = p.attribute("rId").value();
                pugi::xml_document pdoc;
                auto decl = pdoc.append_child(pugi::node_declaration);
                decl.append_attribute("version") = "1.0";
                decl.append_attribute("encoding") = "UTF-8";
                decl.append_attribute("standalone") = "yes";
                pdoc.append_copy(p.first_child());
                std::stringstream ss;
                pdoc.save(ss, "", pugi::format_raw);
                zip_buffers.push_back(ss.str());
                zip_file_add(z, ("word/" + pfile).c_str(), zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0), ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
                std::string relType = (ptype == "header") ? "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
                pimpl->doc_rels.addRelationship(relType, pfile, prid);
                ct += "  <Override PartName=\"/word/" + pfile + "\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml." + ptype + "+xml\"/>\n";
            }
            pimpl->doc.child("w:document").remove_child(parts);
        }

        if (pimpl->doc.child("w:document").child("openword_numbering")) {
            std::string num_xml = 
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\n"
                "  <w:abstractNum w:abstractNumId=\"0\">\n"
                "    <w:nsid w:val=\"FFFFFF01\"/><w:multiLevelType w:val=\"singleLevel\"/>\n"
                "    <w:lvl w:ilvl=\"0\">\n"
                "      <w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"&#xF0B7;\"/><w:lvlJc w:val=\"left\"/>\n"
                "      <w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr>\n"
                "      <w:rPr><w:rFonts w:ascii=\"Symbol\" w:hAnsi=\"Symbol\" w:hint=\"default\"/></w:rPr>\n"
                "    </w:lvl>\n"
                "  </w:abstractNum>\n"
                "  <w:abstractNum w:abstractNumId=\"1\">\n"
                "    <w:nsid w:val=\"FFFFFF02\"/><w:multiLevelType w:val=\"singleLevel\"/>\n"
                "    <w:lvl w:ilvl=\"0\">\n"
                "      <w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"left\"/>\n"
                "      <w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr>\n"
                "    </w:lvl>\n"
                "  </w:abstractNum>\n"
                "  <w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>\n"
                "  <w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/></w:num>\n"
                "</w:numbering>";
            zip_buffers.push_back(num_xml);
            zip_file_add(z, "word/numbering.xml", zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0), ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
            pimpl->doc_rels.addRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "numbering.xml");
            ct += "  <Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>\n";
            pimpl->doc.child("w:document").remove_child("openword_numbering");
        }

        pimpl->doc.child("w:document").remove_attribute("openword_next_rel_id");
        std::stringstream doc_stream;
        pimpl->doc.save(doc_stream, "", pugi::format_raw);
        zip_buffers.push_back(doc_stream.str());
        zip_file_add(z, "word/document.xml", zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0), ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);

        if (!pimpl->doc_rels.empty()) {
            pugi::xml_document rels_doc;
            auto decl3 = rels_doc.append_child(pugi::node_declaration);
            decl3.append_attribute("version") = "1.0";
            decl3.append_attribute("encoding") = "UTF-8";
            decl3.append_attribute("standalone") = "yes";
            auto rels_root = rels_doc.append_child("Relationships");
            rels_root.append_attribute("xmlns") = "http://schemas.openxmlformats.org/package/2006/relationships";
            pimpl->doc_rels.serialize(rels_root);
            std::stringstream rs;
            rels_doc.save(rs, "", pugi::format_raw);
            zip_buffers.push_back(rs.str());
            zip_file_add(z, "word/_rels/document.xml.rels", zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0), ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
        }

        ct += "</Types>";
        zip_file_add(z, "[Content_Types].xml", zip_source_buffer(z, ct.data(), ct.size(), 0), ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);

        std::string root_rels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
            "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>\n"
            "</Relationships>";
        zip_file_add(z, "_rels/.rels", zip_source_buffer(z, root_rels.data(), root_rels.size(), 0), ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);

        zip_close(z);
        return true;
    } catch(...) { return false; }
}

std::string Document::convertMathMLToOMML(const std::string& mathml) const { return convert_mathml_to_omml(mathml); }
std::string Document::convertLaTeXToOMML(const std::string& latex) const { return convert_latex_to_omml(latex); }

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
        if (!read_xml_from_zip("word/document.xml", pimpl->doc)) { zip_close(z); return false; }
        pimpl->body = pimpl->doc.child("w:document").child("w:body");
        if (read_xml_from_zip("word/styles.xml", pimpl->styles_doc)) {
            pimpl->has_styles = true;
            pimpl->styles_root = pimpl->styles_doc.child("w:styles");
        }
        zip_close(z);
        return true;
    } catch (...) { return false; }
}

Section Document::finalSection() { return Section(pimpl->body.child("w:sectPr").internal_object()); }
std::vector<Paragraph> Document::paragraphs() const {
    std::vector<Paragraph> result;
    for (auto node : pimpl->body.children("w:p")) result.push_back(Paragraph(node.internal_object()));
    return result;
}

} // namespace openword
