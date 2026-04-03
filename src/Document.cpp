#include "DefaultStyles.h"
#include "Internal.h"
#include <list>
#include <openword/Document.h>
#include <set>

// Standard library
#include <sstream>

// Third-party libraries
#include <fmt/core.h>
#include <pugixml.hpp>
#include <zip.h>

namespace openword {

// --- Document Implementation ---
struct Document::Impl {
    pugi::xml_document doc;
    pugi::xml_node body;
    pugi::xml_document styles_doc;
    pugi::xml_node styles_root;
    RelationshipManager doc_rels;
    bool has_styles = false;
    pugi::xml_document numbering_doc;
    bool has_numbering = false;

    bool even_and_odd_headers = false;

    Metadata metadata;
    bool has_metadata = false;

    pugi::xml_document footnotes_doc;
    pugi::xml_node footnotes_root;
    int next_footnote_id = 1;

    pugi::xml_document endnotes_doc;
    pugi::xml_node endnotes_root;
    int next_endnote_id = 1;

    pugi::xml_document comments_doc;
    pugi::xml_node comments_root;
    int next_comment_id = 0;

    bool needs_update_fields = false;

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

        styles_doc.load_string(DEFAULT_STYLES_XML);
        styles_root = styles_doc.child("w:styles");
        has_styles = true;

        auto f_decl = footnotes_doc.append_child(pugi::node_declaration);
        f_decl.append_attribute("version") = "1.0";
        f_decl.append_attribute("encoding") = "UTF-8";
        f_decl.append_attribute("standalone") = "yes";
        footnotes_root = footnotes_doc.append_child("w:footnotes");
        footnotes_root.append_attribute("xmlns:w") = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        auto e_decl = endnotes_doc.append_child(pugi::node_declaration);
        e_decl.append_attribute("version") = "1.0";
        e_decl.append_attribute("encoding") = "UTF-8";
        e_decl.append_attribute("standalone") = "yes";
        endnotes_root = endnotes_doc.append_child("w:endnotes");
        endnotes_root.append_attribute("xmlns:w") = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        auto c_decl = comments_doc.append_child(pugi::node_declaration);
        c_decl.append_attribute("version") = "1.0";
        c_decl.append_attribute("encoding") = "UTF-8";
        c_decl.append_attribute("standalone") = "yes";
        comments_root = comments_doc.append_child("w:comments");
        comments_root.append_attribute("xmlns:w") = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    }
};

Document::Document() : pimpl(std::make_unique<Impl>()) { pimpl->initialize(); }

Document::Document(const std::string &templatePath) : pimpl(std::make_unique<Impl>()) {
    pimpl->initialize(); // Set up basic empty doc structure

    int error = 0;
    zip_t *z = zip_open(templatePath.c_str(), ZIP_RDONLY, &error);
    if (z) {
        auto read_xml_from_zip = [&](const char *filename, pugi::xml_document &target_doc) -> bool {
            zip_stat_t st;
            zip_stat_init(&st);
            if (zip_stat(z, filename, 0, &st) != 0)
                return false;
            zip_file_t *f = zip_fopen(z, filename, 0);
            if (!f)
                return false;
            std::string file_content(st.size, '\0');
            zip_fread(f, file_content.data(), st.size);
            zip_fclose(f);
            return target_doc.load_string(file_content.c_str());
        };

        // Extract styles and numbering from the template to override the defaults
        if (read_xml_from_zip("word/styles.xml", pimpl->styles_doc)) {
            pimpl->styles_root = pimpl->styles_doc.child("w:styles");
            pimpl->has_styles = true;
        }

        if (read_xml_from_zip("word/numbering.xml", pimpl->numbering_doc)) {
            pimpl->has_numbering = true;
        }
        zip_close(z);
    }
}
Document::~Document() = default;

void Document::setEvenAndOddHeaders(bool val) { pimpl->even_and_odd_headers = val; }

Paragraph Document::addParagraph(const std::string &text) {
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
    for (int c = 0; c < cols; ++c)
        tblGrid.append_child("w:gridCol");
    for (int r = 0; r < rows; ++r) {
        auto tr = tbl.append_child("w:tr");
        for (int c = 0; c < cols; ++c) {
            auto tc = tr.append_child("w:tc");
            auto tcPr = tc.append_child("w:tcPr");
            auto tcW = tcPr.append_child("w:tcW");
            tcW.append_attribute("w:w") = "0";
            tcW.append_attribute("w:type") = "auto";

            // Critical OOXML Requirement: Every w:tc must end with at least one paragraph (even an empty one)
            tc.append_child("w:p");
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
        zip_t *z = zip_open(filepath, ZIP_CREATE | ZIP_TRUNCATE, &error);
        if (!z)
            return false;

        std::list<std::string> zip_buffers;
        auto safe_zip_add = [&](const char *name, zip_source_t *source, zip_flags_t flags) {
            if (!source)
                return;
            if (zip_file_add(z, name, source, flags) < 0) {
                zip_source_free(source);
            }
        };

        int start_id = pimpl->doc.child("w:document").attribute("openword_next_rel_id").as_int(100);
        pimpl->doc_rels.setNextId(start_id);

        std::string ct =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
            "  <Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>\n"
            "  <Default Extension=\"png\" ContentType=\"image/png\"/>\n"
            "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
            "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
            "  <Override PartName=\"/word/document.xml\" "
            "ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\n";

        if (pimpl->has_metadata) {
            std::string core_props =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                "<cp:coreProperties "
                "xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
                "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" "
                "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" "
                "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n";

            if (!pimpl->metadata.title.empty())
                core_props += fmt::format("  <dc:title>{}</dc:title>\n", pimpl->metadata.title);
            if (!pimpl->metadata.subject.empty())
                core_props += fmt::format("  <dc:subject>{}</dc:subject>\n", pimpl->metadata.subject);
            if (!pimpl->metadata.author.empty())
                core_props += fmt::format("  <dc:creator>{}</dc:creator>\n", pimpl->metadata.author);
            if (!pimpl->metadata.keywords.empty())
                core_props += fmt::format("  <cp:keywords>{}</cp:keywords>\n", pimpl->metadata.keywords);
            if (!pimpl->metadata.comments.empty())
                core_props += fmt::format("  <dc:description>{}</dc:description>\n", pimpl->metadata.comments);
            if (!pimpl->metadata.lastModifiedBy.empty())
                core_props +=
                    fmt::format("  <cp:lastModifiedBy>{}</cp:lastModifiedBy>\n", pimpl->metadata.lastModifiedBy);
            if (!pimpl->metadata.category.empty())
                core_props += fmt::format("  <cp:category>{}</cp:category>\n", pimpl->metadata.category);

            core_props += "</cp:coreProperties>";

            zip_buffers.push_back(core_props);
            safe_zip_add("docProps/core.xml",
                         zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0),
                         ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
            ct += "  <Override PartName=\"/docProps/core.xml\" "
                  "ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\n";

            // App.xml (Extended Properties)
            std::string app_props =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" "
                "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\n"
                "  <Application>OpenWord C++</Application>\n"
                "  <DocSecurity>0</DocSecurity>\n"
                "  <ScaleCrop>false</ScaleCrop>\n"
                "  <HeadingPairs><vt:vector size=\"2\" "
                "baseType=\"variant\"><vt:variant><vt:lpstr>Title</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></"
                "vt:variant></vt:vector></HeadingPairs>\n"
                "  <TitlesOfParts><vt:vector size=\"1\" baseType=\"lpstr\"><vt:lpstr>" +
                (pimpl->metadata.title.empty() ? "Document" : pimpl->metadata.title) +
                "</vt:lpstr></vt:vector></TitlesOfParts>\n"
                "  <LinksUpToDate>false</LinksUpToDate>\n"
                "  <SharedDoc>false</SharedDoc>\n"
                "  <HyperlinksChanged>false</HyperlinksChanged>\n"
                "  <AppVersion>16.0000</AppVersion>\n";

            if (!pimpl->metadata.company.empty())
                app_props += fmt::format("  <Company>{}</Company>\n", pimpl->metadata.company);
            if (!pimpl->metadata.manager.empty())
                app_props += fmt::format("  <Manager>{}</Manager>\n", pimpl->metadata.manager);
            if (!pimpl->metadata.hyperlinkBase.empty())
                app_props += fmt::format("  <HyperlinkBase>{}</HyperlinkBase>\n", pimpl->metadata.hyperlinkBase);

            app_props += "</Properties>";
            zip_buffers.push_back(app_props);
            safe_zip_add("docProps/app.xml",
                         zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0),
                         ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
            ct += "  <Override PartName=\"/docProps/app.xml\" "
                  "ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>\n";

            // Custom.xml (Custom Properties)
            if (!pimpl->metadata.customProperties.empty()) {
                std::string custom_props =
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                    "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" "
                    "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\n";

                int pid = 2;
                for (const auto &[name, prop] : pimpl->metadata.customProperties) {
                    custom_props += fmt::format(
                        "  <property fmtid=\"{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}\" pid=\"{}\" name=\"{}\">\n",
                        pid++, name);

                    if (prop.type == CustomProperty::Type::Text) {
                        custom_props += fmt::format("    <vt:lpwstr>{}</vt:lpwstr>\n", prop.value);
                    } else if (prop.type == CustomProperty::Type::Integer) {
                        custom_props += fmt::format("    <vt:i4>{}</vt:i4>\n", prop.value);
                    } else if (prop.type == CustomProperty::Type::Boolean) {
                        custom_props += fmt::format("    <vt:bool>{}</vt:bool>\n", prop.value);
                    } else if (prop.type == CustomProperty::Type::Double) {
                        custom_props += fmt::format("    <vt:r8>{}</vt:r8>\n", prop.value);
                    } else if (prop.type == CustomProperty::Type::Date) {
                        custom_props += fmt::format("    <vt:filetime>{}</vt:filetime>\n", prop.value);
                    }

                    custom_props += "  </property>\n";
                }
                custom_props += "</Properties>";

                zip_buffers.push_back(custom_props);
                safe_zip_add("docProps/custom.xml",
                             zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0),
                             ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
                ct += "  <Override PartName=\"/docProps/custom.xml\" "
                      "ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/>\n";
            }
        }

        if (pimpl->next_footnote_id > 1) {
            std::stringstream fs;
            pimpl->footnotes_doc.save(fs, "", pugi::format_raw);
            zip_buffers.push_back(fs.str());
            safe_zip_add("word/footnotes.xml",
                         zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0),
                         ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
            pimpl->doc_rels.addRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes", "footnotes.xml");
            ct += "  <Override PartName=\"/word/footnotes.xml\" "
                  "ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml\"/>\n";
        }

        if (pimpl->next_endnote_id > 1) {
            std::stringstream es;
            pimpl->endnotes_doc.save(es, "", pugi::format_raw);
            zip_buffers.push_back(es.str());
            safe_zip_add("word/endnotes.xml",
                         zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0),
                         ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
            pimpl->doc_rels.addRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes", "endnotes.xml");
            ct += "  <Override PartName=\"/word/endnotes.xml\" "
                  "ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml\"/>\n";
        }

        if (pimpl->next_comment_id > 0) {
            std::stringstream cs;
            pimpl->comments_doc.save(cs, "", pugi::format_raw);
            zip_buffers.push_back(cs.str());
            safe_zip_add("word/comments.xml",
                         zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0),
                         ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
            pimpl->doc_rels.addRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", "comments.xml");
            ct += "  <Override PartName=\"/word/comments.xml\" "
                  "ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml\"/>\n";
        }

        auto links_store = pimpl->doc.child("w:document").child("openword_hyperlinks");
        if (links_store) {
            for (auto m : links_store.children("link")) {
                std::string url = m.attribute("url").value();
                std::string rid = m.attribute("rId").value();
                pimpl->doc_rels.addRelationship(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", url, rid,
                    "External");
            }
            pimpl->doc.child("w:document").remove_child(links_store);
        }

        if (pimpl->has_styles) {
            std::stringstream ss;
            pimpl->styles_doc.save(ss, "", pugi::format_raw);
            zip_buffers.push_back(ss.str());
            safe_zip_add("word/styles.xml",
                         zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0),
                         ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
            pimpl->doc_rels.addRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml");
            ct += "  <Override PartName=\"/word/styles.xml\" "
                  "ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>\n";
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
                if (dot_pos != std::string::npos)
                    ext = path.substr(dot_pos + 1);
                std::string target = "media/image" + std::to_string(media_count) + "." + ext;
                pimpl->doc_rels.addRelationship(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target, old_rid);
                safe_zip_add(("word/" + target).c_str(), zip_source_file(z, path.c_str(), 0, 0), ZIP_FL_OVERWRITE);
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
                safe_zip_add(("word/" + pfile).c_str(),
                             zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0),
                             ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
                std::string relType =
                    (ptype == "header") ? "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
                                        : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
                pimpl->doc_rels.addRelationship(relType, pfile, prid);
                ct += "  <Override PartName=\"/word/" + pfile +
                      "\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml." + ptype +
                      "+xml\"/>\n";
            }
            pimpl->doc.child("w:document").remove_child(parts);
        }

        if (pimpl->has_numbering) {
            std::stringstream ss;
            pimpl->numbering_doc.save(ss, "", pugi::format_raw);
            zip_buffers.push_back(ss.str());
            safe_zip_add("word/numbering.xml",
                         zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0),
                         ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
            pimpl->doc_rels.addRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "numbering.xml");
            ct += "  <Override PartName=\"/word/numbering.xml\" "
                  "ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>\n";
        }

        int current_rel_id = pimpl->doc.child("w:document").attribute("openword_next_rel_id").as_int(100);
        pimpl->doc.child("w:document").remove_attribute("openword_next_rel_id");
        std::stringstream doc_stream;
        pimpl->doc.save(doc_stream, "", pugi::format_raw);
        zip_buffers.push_back(doc_stream.str());
        safe_zip_add("word/document.xml", zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0),
                     ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
        pimpl->doc.child("w:document").append_attribute("openword_next_rel_id") =
            std::to_string(current_rel_id).c_str();

        if (pimpl->needs_update_fields || pimpl->even_and_odd_headers) {
            std::string settings_xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\n";
            if (pimpl->needs_update_fields) {
                settings_xml += "  <w:updateFields w:val=\"true\"/>\n";
            }
            if (pimpl->even_and_odd_headers) {
                settings_xml += "  <w:evenAndOddHeaders/>\n";
            }
            settings_xml += "</w:settings>";
            zip_buffers.push_back(settings_xml);
            safe_zip_add("word/settings.xml",
                         zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0),
                         ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
            pimpl->doc_rels.addRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings", "settings.xml");
            ct += "  <Override PartName=\"/word/settings.xml\" "
                  "ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>\n";
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
            std::stringstream rs;
            rels_doc.save(rs, "", pugi::format_raw);
            zip_buffers.push_back(rs.str());
            safe_zip_add("word/_rels/document.xml.rels",
                         zip_source_buffer(z, zip_buffers.back().data(), zip_buffers.back().size(), 0),
                         ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
        }

        ct += "</Types>";
        safe_zip_add("[Content_Types].xml", zip_source_buffer(z, ct.data(), ct.size(), 0),
                     ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);

        std::string root_rels =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
            "  <Relationship Id=\"rId1\" "
            "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
            "Target=\"word/document.xml\"/>\n";
        if (pimpl->has_metadata) {
            root_rels += "  <Relationship Id=\"rId2\" "
                         "Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/"
                         "core-properties\" Target=\"docProps/core.xml\"/>\n";
            root_rels += "  <Relationship Id=\"rId3\" "
                         "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                         "extended-properties\" Target=\"docProps/app.xml\"/>\n";
            if (!pimpl->metadata.customProperties.empty()) {
                root_rels += "  <Relationship Id=\"rId4\" "
                             "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                             "custom-properties\" Target=\"docProps/custom.xml\"/>\n";
            }
        }
        root_rels += "</Relationships>";
        safe_zip_add("_rels/.rels", zip_source_buffer(z, root_rels.data(), root_rels.size(), 0),
                     ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);

        zip_close(z);
        return true;
    } catch (...) {
        return false;
    }
}

std::string Document::convertMathMLToOMML(const std::string &mathml) const { return convert_mathml_to_omml(mathml); }
std::string Document::convertLaTeXToOMML(const std::string &latex) const { return convert_latex_to_omml(latex); }

void Document::setMetadata(const Metadata &meta) {
    pimpl->metadata = meta;
    pimpl->has_metadata = true;
}

Metadata Document::metadata() const { return pimpl->metadata; }

int Document::createFootnote(const std::string &text) {
    int id = pimpl->next_footnote_id++;
    auto fn = pimpl->footnotes_root.append_child("w:footnote");
    fn.append_attribute("w:id") = std::to_string(id).c_str();
    auto p = fn.append_child("w:p");
    auto pPr = p.append_child("w:pPr");
    pPr.append_child("w:pStyle").append_attribute("w:val") = "FootnoteText";
    auto r1 = p.append_child("w:r");
    r1.append_child("w:rPr").append_child("w:rStyle").append_attribute("w:val") = "FootnoteReference";
    r1.append_child("w:footnoteRef");
    auto rSpace = p.append_child("w:r");
    rSpace.append_child("w:t").text().set(" ");
    auto r2 = p.append_child("w:r");
    r2.append_child("w:t").text().set(text.c_str());
    return id;
}

int Document::createEndnote(const std::string &text) {
    int id = pimpl->next_endnote_id++;
    auto en = pimpl->endnotes_root.append_child("w:endnote");
    en.append_attribute("w:id") = std::to_string(id).c_str();
    auto p = en.append_child("w:p");
    auto pPr = p.append_child("w:pPr");
    pPr.append_child("w:pStyle").append_attribute("w:val") = "EndnoteText";
    auto r1 = p.append_child("w:r");
    r1.append_child("w:rPr").append_child("w:rStyle").append_attribute("w:val") = "EndnoteReference";
    r1.append_child("w:endnoteRef");
    auto rSpace = p.append_child("w:r");
    rSpace.append_child("w:t").text().set(" ");
    auto r2 = p.append_child("w:r");
    r2.append_child("w:t").text().set(text.c_str());
    return id;
}

StyleCollection Document::styles() { return StyleCollection(pimpl->styles_root.internal_object()); }

NumberingCollection Document::numbering() {
    pimpl->has_numbering = true;
    if (!pimpl->numbering_doc.child("w:numbering")) {
        auto decl = pimpl->numbering_doc.append_child(pugi::node_declaration);
        decl.append_attribute("version") = "1.0";
        decl.append_attribute("encoding") = "UTF-8";
        decl.append_attribute("standalone") = "yes";
        auto root = pimpl->numbering_doc.append_child("w:numbering");
        root.append_attribute("xmlns:w") = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    }
    return NumberingCollection(pimpl->numbering_doc.child("w:numbering").internal_object());
}

void Document::addTableOfContents(gsl::czstring title, int max_levels, TOCLeader leader) {
    pimpl->needs_update_fields = true;
    auto sectPr = pimpl->body.child("w:sectPr");

    if (title && title[0] != '\0') {
        auto pTitle = pimpl->body.insert_child_before("w:p", sectPr);
        pTitle.append_child("w:pPr").append_child("w:pStyle").append_attribute("w:val") = "TOCHeading";
        pTitle.append_child("w:r").append_child("w:t").text().set(title);
    }

    auto sdt = pimpl->body.insert_child_before("w:sdt", sectPr);
    auto sdtPr = sdt.append_child("w:sdtPr");
    auto docPartObj = sdtPr.append_child("w:docPartObj");
    docPartObj.append_child("w:docPartGallery").append_attribute("w:val") = "Table of Contents";
    docPartObj.append_child("w:docPartUnique");
    auto sdtContent = sdt.append_child("w:sdtContent");

    auto p1 = sdtContent.append_child("w:p");
    p1.append_child("w:pPr").append_child("w:pStyle").append_attribute("w:val") = "TOC1";

    auto r1 = p1.append_child("w:r");
    auto fldBegin = r1.append_child("w:fldChar");
    fldBegin.append_attribute("w:fldCharType") = "begin";
    fldBegin.append_attribute("w:dirty") = "true";

    auto r2 = p1.append_child("w:r");
    auto instr = r2.append_child("w:instrText");
    instr.append_attribute("xml:space") = "preserve";

    std::string leaderChar = "";
    if (leader == TOCLeader::Hyphen)
        leaderChar = " \\p \"-\"";
    else if (leader == TOCLeader::Underscore)
        leaderChar = " \\p \"_\"";
    else if (leader == TOCLeader::None)
        leaderChar = " \\p \" \"";
    // default dot doesn't need explicit \p switch

    std::string instrVal = fmt::format(R"(TOC \o "1-{}" \h \z \u{})", max_levels, leaderChar);
    instr.text().set(instrVal.c_str());

    auto r3 = p1.append_child("w:r");
    r3.append_child("w:fldChar").append_attribute("w:fldCharType") = "separate";

    auto p2 = sdtContent.append_child("w:p");
    p2.append_child("w:pPr").append_child("w:pStyle").append_attribute("w:val") = "TOC1";
    p2.append_child("w:r").append_child("w:t").text().set(
        "Right-click here and select 'Update Field' to generate TOC.");

    auto p3 = sdtContent.append_child("w:p");
    p3.append_child("w:pPr").append_child("w:pStyle").append_attribute("w:val") = "TOC1";
    p3.append_child("w:r").append_child("w:fldChar").append_attribute("w:fldCharType") = "end";
}

bool Document::load(gsl::czstring filepath) {
    try {
        int error = 0;
        zip_t *z = zip_open(filepath, ZIP_RDONLY, &error);
        if (!z)
            return false;
        pimpl->doc.reset();
        pimpl->styles_doc.reset();
        pimpl->has_styles = false;
        auto read_xml_from_zip = [&](const char *filename, pugi::xml_document &target_doc) -> bool {
            zip_stat_t st;
            zip_stat_init(&st);
            if (zip_stat(z, filename, 0, &st) != 0)
                return false;
            zip_file_t *f = zip_fopen(z, filename, 0);
            if (!f)
                return false;
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

Section Document::finalSection() { return Section(pimpl->body.child("w:sectPr").internal_object()); }
std::vector<Paragraph> Document::paragraphs() const {
    std::vector<Paragraph> result;
    for (auto node : pimpl->body.children("w:p"))
        result.push_back(Paragraph(node.internal_object()));
    return result;
}

std::vector<Table> Document::tables() const {
    std::vector<Table> result;
    for (auto node : pimpl->body.children("w:tbl"))
        result.push_back(Table(node.internal_object()));
    return result;
}

int Document::replaceText(const std::string &search, const std::string &replace) {
    int count = 0;

    // Replace in body paragraphs and tables
    for (auto node : pimpl->body.children()) {
        std::string name = node.name();
        if (name == "w:p") {
            count += Paragraph(node.internal_object()).replaceText(search, replace);
        } else if (name == "w:tbl") {
            count += Table(node.internal_object()).replaceText(search, replace);
        }
    }

    // Replace in headers and footers
    auto parts = pimpl->doc.child("w:document").child("openword_parts");
    if (parts) {
        for (auto part : parts.children("part")) {
            for (auto rootNode : part.children()) {
                std::string rootName = rootNode.name();
                if (rootName == "w:hdr" || rootName == "w:ftr") {
                    for (auto node : rootNode.children()) {
                        std::string name = node.name();
                        if (name == "w:p") {
                            count += Paragraph(node.internal_object()).replaceText(search, replace);
                        } else if (name == "w:tbl") {
                            count += Table(node.internal_object()).replaceText(search, replace);
                        }
                    }
                }
            }
        }
    }

    return count;
}

int Document::cloneRowAndSetValues(const std::string &search,
                                   const std::vector<std::map<std::string, std::string>> &values) {
    if (search.empty())
        return 0;

    struct TextExtractor : pugi::xml_tree_walker {
        std::string text;
        bool for_each(pugi::xml_node &node) override {
            if (std::strcmp(node.name(), "w:t") == 0) {
                text += node.text().get();
            }
            return true;
        }
    };

    auto process_rows = [&](pugi::xml_node root) -> int {
        int created_count = 0;
        std::vector<pugi::xml_node> target_trs;

        struct TrCollector : pugi::xml_tree_walker {
            const std::string &search_str;
            std::vector<pugi::xml_node> &trs;
            TrCollector(const std::string &s, std::vector<pugi::xml_node> &t) : search_str(s), trs(t) {}

            bool for_each(pugi::xml_node &node) override {
                if (std::strcmp(node.name(), "w:tr") == 0) {
                    TextExtractor ext;
                    node.traverse(ext);
                    if (ext.text.find(search_str) != std::string::npos) {
                        trs.push_back(node);
                    }
                }
                return true;
            }
        };

        TrCollector collector(search, target_trs);
        root.traverse(collector);

        for (auto tr : target_trs) {
            auto parent = tr.parent(); // <w:tbl>
            if (!parent)
                continue;

            for (const auto &row_map : values) {
                pugi::xml_node new_tr = parent.insert_copy_before(tr, tr);
                Row r(new_tr.internal_object());
                for (const auto &kv : row_map) {
                    r.replaceText(kv.first, kv.second);
                }
                created_count++;
            }
            // Remove the template row
            parent.remove_child(tr);
        }
        return created_count;
    };

    int count = process_rows(pimpl->body);

    auto parts = pimpl->doc.child("w:document").child("openword_parts");
    if (parts) {
        for (auto part : parts.children("part")) {
            for (auto rootNode : part.children()) {
                std::string rootName = rootNode.name();
                if (rootName == "w:hdr" || rootName == "w:ftr") {
                    count += process_rows(rootNode);
                }
            }
        }
    }
    return count;
}

std::vector<BlockElement> Document::elements() const {
    std::vector<BlockElement> result;
    for (auto node : pimpl->body.children()) {
        std::string name = node.name();
        if (name == "w:p") {
            result.push_back(Paragraph(node.internal_object()));
        } else if (name == "w:tbl") {
            result.push_back(Table(node.internal_object()));
        } else if (name == "w:sectPr") {
            result.push_back(Section(node.internal_object()));
        }
    }
    return result;
}

void Document::addWatermark(const std::string &text) {
    if (text.empty())
        return;

    auto sec = finalSection();
    auto header = sec.addHeader(HeaderFooterType::Default);
    auto p = header.addParagraph();

    // We do NOT modify paragraph spacing or size. We leave it as a normal header paragraph.
    // The previous attempt to hide it with spacing=0 probably caused Word to clip or ignore the VML shape entirely.

    std::string vmlStr = fmt::format(R"(
        <w:r>
            <w:pict>
                <v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">
                    <v:formulas>
                        <v:f eqn="sum #0 0 10800"/>
                        <v:f eqn="prod #0 2 1"/>
                        <v:f eqn="sum 21600 0 @1"/>
                        <v:f eqn="sum 0 0 @2"/>
                        <v:f eqn="sum 21600 0 @3"/>
                        <v:f eqn="if @0 @3 0"/>
                        <v:f eqn="if @0 21600 @1"/>
                        <v:f eqn="if @0 0 @2"/>
                        <v:f eqn="if @0 @4 21600"/>
                        <v:f eqn="mid @5 @6"/>
                        <v:f eqn="mid @8 @5"/>
                        <v:f eqn="mid @7 @8"/>
                        <v:f eqn="mid @6 @7"/>
                        <v:f eqn="sum @6 0 @5"/>
                    </v:formulas>
                    <v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>
                    <v:textpath on="t" fitshape="t"/>
                    <v:handles>
                        <v:h position="#0,bottomRight" xrange="6629,14971"/>
                    </v:handles>
                    <o:lock v:ext="edit" text="t" shapetype="t"/>
                </v:shapetype>
                <v:shape id="PowerPlusWaterMarkObject357" o:spid="_x0000_s102" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:415pt;height:90pt;rotation:315;z-index:-251658240;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" fillcolor="#c0c0c0" stroked="f">
                    <v:fill opacity="0.5"/>
                    <v:textpath style="font-family:'Calibri';font-size:1pt" string="{}"/>
                </v:shape>
            </w:pict>
        </w:r>
    )",
                                     text);

    p.addRawXml(vmlStr);
}


void Document::addHtml(const std::string &html) {
    parseHtmlAndInsert(html, [this]() { return this->addParagraph(); });
}

} // namespace openword

int openword::Document::createComment(const std::string &text, const std::string &author, const std::string &initials) {
    int id = pimpl->next_comment_id++;
    auto cmt = pimpl->comments_root.append_child("w:comment");
    cmt.append_attribute("w:id") = std::to_string(id).c_str();
    if (!author.empty())
        cmt.append_attribute("w:author") = author.c_str();
    if (!initials.empty())
        cmt.append_attribute("w:initials") = initials.c_str();

    // Add current time in UTC using standard library formatting (if available) or rely on Word interpreting missing
    // time. Let's omit date for simplicity and testing consistency.

    auto p = cmt.append_child("w:p");
    auto pPr = p.append_child("w:pPr");
    pPr.append_child("w:pStyle").append_attribute("w:val") = "CommentText";
    auto r1 = p.append_child("w:r");
    r1.append_child("w:rPr").append_child("w:rStyle").append_attribute("w:val") = "CommentReference";
    r1.append_child("w:annotationRef"); // OOXML requirement for comments
    auto r2 = p.append_child("w:r");
    r2.append_child("w:t").text().set(text.c_str());
    return id;
}

bool openword::Document::validate(gsl::czstring partName, const openword::SchemaValidator &validator,
                                  std::string &outErrors) const {
    if (!validator.isValid()) {
        outErrors = "Invalid schema validator.";
        return false;
    }

    std::string name(partName);
    std::stringstream stream;

    if (name == "word/document.xml" || name == "document.xml") {
        pimpl->doc.save(stream, "", pugi::format_raw);
    } else if (name == "word/styles.xml" || name == "styles.xml") {
        if (!pimpl->has_styles)
            return true; // valid if not present
        pimpl->styles_doc.save(stream, "", pugi::format_raw);
    } else if (name == "word/numbering.xml" || name == "numbering.xml") {
        if (!pimpl->has_numbering)
            return true;
        pimpl->numbering_doc.save(stream, "", pugi::format_raw);
    } else if (name == "word/footnotes.xml" || name == "footnotes.xml") {
        pimpl->footnotes_doc.save(stream, "", pugi::format_raw);
    } else if (name == "word/endnotes.xml" || name == "endnotes.xml") {
        pimpl->endnotes_doc.save(stream, "", pugi::format_raw);
    } else if (name == "word/comments.xml" || name == "comments.xml") {
        pimpl->comments_doc.save(stream, "", pugi::format_raw);
    } else {
        outErrors = "Unknown or unsupported part name for validation: " + name;
        return false;
    }

    std::string xmlContent = stream.str();
    return validator.validate(xmlContent.c_str(), outErrors);
}
