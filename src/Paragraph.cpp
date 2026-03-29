#include <openword/Document.h>
#include "Internal.h"

#include <pugixml.hpp>
#include <fstream>

namespace openword {

// Returns {width, height} in pixels
static std::pair<int, int> getImageSize(const std::string& path) {
    std::ifstream file;
    open_ifstream(file, path, std::ios::binary);
    if (!file.is_open()) return {0, 0};

    unsigned char buf[24];
    if (file.read(reinterpret_cast<char*>(buf), 24)) {
        // Check PNG
        if (buf[0] == 0x89 && buf[1] == 0x50 && buf[2] == 0x4E && buf[3] == 0x47) {
            int w = (buf[16] << 24) | (buf[17] << 16) | (buf[18] << 8) | buf[19];
            int h = (buf[20] << 24) | (buf[21] << 16) | (buf[22] << 8) | buf[23];
            return {w, h};
        }
        
        // Check JPEG (naive marker parsing)
        if (buf[0] == 0xFF && buf[1] == 0xD8) {
            file.seekg(2, std::ios::beg);
            while (true) {
                int marker = file.get();
                if (marker != 0xFF) break;
                int type = file.get();
                if (type == EOF || file.eof()) break;
                
                int len_hi = file.get();
                int len_lo = file.get();
                int len = (len_hi << 8) | len_lo;
                
                // SOF0 (0xC0) to SOF15 (0xCF) except DHT (0xC4), JPG (0xC8), DAC (0xCC)
                if (type >= 0xC0 && type <= 0xCF && type != 0xC4 && type != 0xC8 && type != 0xCC) {
                    file.get(); // precision
                    int h_hi = file.get();
                    int h_lo = file.get();
                    int w_hi = file.get();
                    int w_lo = file.get();
                    return {(w_hi << 8) | w_lo, (h_hi << 8) | h_lo};
                } else {
                    file.seekg(len - 2, std::ios::cur);
                }
            }
        }
    }
    return {0, 0};
}

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

Paragraph& Paragraph::setOutlineLevel(int level) {
    auto n = cast_node(node_);
    auto pPr = n.child("w:pPr");
    if (!pPr) {
        pPr = n.prepend_child("w:pPr");
    }
    auto outline = pPr.child("w:outlineLvl");
    if (!outline) {
        outline = pPr.append_child("w:outlineLvl");
    }
    outline.remove_attribute("w:val");
    outline.append_attribute("w:val").set_value(std::to_string(level).c_str());
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

Paragraph& Paragraph::setSpacing(int beforeTwips, int afterTwips, int lineSpacing, gsl::czstring lineRule) {
    auto n = cast_node(node_);
    auto pPr = n.child("w:pPr");
    if (!pPr) pPr = n.prepend_child("w:pPr");
    auto spacing = pPr.child("w:spacing");
    if (!spacing) spacing = pPr.append_child("w:spacing");
    
    spacing.remove_attribute("w:before");
    spacing.remove_attribute("w:after");
    spacing.remove_attribute("w:line");
    spacing.remove_attribute("w:lineRule");

    spacing.append_attribute("w:before").set_value(std::to_string(beforeTwips).c_str());
    spacing.append_attribute("w:after").set_value(std::to_string(afterTwips).c_str());
    if (lineSpacing > 0) {
        spacing.append_attribute("w:line").set_value(std::to_string(lineSpacing).c_str());
        spacing.append_attribute("w:lineRule").set_value(lineRule);
    }
    return *this;
}

Paragraph& Paragraph::setIndentation(int leftTwips, int rightTwips, int firstLineTwips, int hangingTwips) {
    auto n = cast_node(node_);
    auto pPr = n.child("w:pPr");
    if (!pPr) pPr = n.prepend_child("w:pPr");
    auto ind = pPr.child("w:ind");
    if (!ind) ind = pPr.append_child("w:ind");

    ind.remove_attribute("w:left");
    ind.remove_attribute("w:right");
    ind.remove_attribute("w:firstLine");
    ind.remove_attribute("w:hanging");

    ind.append_attribute("w:left").set_value(std::to_string(leftTwips).c_str());
    ind.append_attribute("w:right").set_value(std::to_string(rightTwips).c_str());
    if (firstLineTwips > 0) ind.append_attribute("w:firstLine").set_value(std::to_string(firstLineTwips).c_str());
    if (hangingTwips > 0) ind.append_attribute("w:hanging").set_value(std::to_string(hangingTwips).c_str());
    return *this;
}

Paragraph& Paragraph::setList(int numIdVal, int level) {
    auto n = cast_node(node_);
    auto pPr = n.child("w:pPr");
    if (!pPr) pPr = n.prepend_child("w:pPr");
    auto numPr = pPr.child("w:numPr");
    
    if (numIdVal <= 0) {
        if (numPr) pPr.remove_child(numPr);
        return *this;
    }
    
    if (!numPr) numPr = pPr.append_child("w:numPr");
    
    auto ilvl = numPr.child("w:ilvl");
    if (!ilvl) ilvl = numPr.append_child("w:ilvl");
    ilvl.remove_attribute("w:val");
    ilvl.append_attribute("w:val") = std::to_string(level).c_str();
    
    auto numId = numPr.child("w:numId");
    if (!numId) numId = numPr.append_child("w:numId");
    numId.remove_attribute("w:val");
    numId.append_attribute("w:val") = std::to_string(numIdVal).c_str();
    
    return *this;
}

Section Paragraph::appendSectionBreak() {
    auto n = cast_node(node_);
    auto body = n.parent();
    auto finalSectPr = body.child("w:sectPr");
    
    // Create pPr if not exists
    auto pPr = n.child("w:pPr");
    if (!pPr) pPr = n.prepend_child("w:pPr");
    
    // Create sectPr in current paragraph
    auto pSectPr = pPr.child("w:sectPr");
    if (!pSectPr) pSectPr = pPr.append_child("w:sectPr");
    else pSectPr.remove_children(); // Ensure clean slate
    
    // If there's a global section at the end, move its content to this break
    if (finalSectPr) {
        // Copy all children of global section to this paragraph's section break
        for (auto child : finalSectPr.children()) {
            pSectPr.append_copy(child);
        }
        // Also copy attributes
        for (auto attr : finalSectPr.attributes()) {
            pSectPr.append_attribute(attr.name()) = attr.value();
        }
        
        // Now reset the global section to defaults for the next section
        finalSectPr.remove_children();
        for (auto attr : finalSectPr.attributes()) {
            finalSectPr.remove_attribute(attr);
        }
        
        // Re-add basic default properties to the new global section
        auto pgSz = finalSectPr.append_child("w:pgSz");
        pgSz.append_attribute("w:w") = "11906";
        pgSz.append_attribute("w:h") = "16838";
        auto pgMar = finalSectPr.append_child("w:pgMar");
        pgMar.append_attribute("w:top") = "1440";
        pgMar.append_attribute("w:right") = "1440";
        pgMar.append_attribute("w:bottom") = "1440";
        pgMar.append_attribute("w:left") = "1440";
        pgMar.append_attribute("w:header") = "720";
        pgMar.append_attribute("w:footer") = "720";
        pgMar.append_attribute("w:gutter") = "0";
    }
    
    return Section(finalSectPr.internal_object());
}

std::vector<Run> Paragraph::runs() const {
    std::vector<Run> result;
    auto n = cast_node(node_);
    for (auto node : n.children("w:r")) {
        result.push_back(Run(node.internal_object()));
    }
    return result;
}

std::string Paragraph::text() const {
    std::string result;
    for (const auto& r : runs()) {
        result += r.text();
    }
    return result;
}

void Paragraph::addImage(gsl::czstring image_path, double scale, ImagePosition position, long long xOffset, long long yOffset) {
    auto n = cast_node(node_);
    auto root = n.root().child("w:document");
    auto media_store = root.child("openword_media");
    if (!media_store) {
        media_store = root.append_child("openword_media");
    }

    int c = 0;
    std::string rid;
    // Check for resource pooling (deduplication by path)
    for (auto child : media_store.children("media")) {
        c++;
        if (std::string(child.attribute("path").value()) == image_path) {
            rid = child.attribute("rId").value();
        }
    }
    
    // If not found in pool, add it
    if (rid.empty()) {
        rid = "rIdMedia" + std::to_string(c + 1);
        auto media_entry = media_store.append_child("media");
        media_entry.append_attribute("path").set_value(image_path);
        media_entry.append_attribute("rId").set_value(rid.c_str());
    }

    auto r = n.append_child("w:r");
    auto drawing = r.append_child("w:drawing");
    
    pugi::xml_node wrapper_node;
    if (position == ImagePosition::Inline) {
        wrapper_node = drawing.append_child("wp:inline");
    } else {
        wrapper_node = drawing.append_child("wp:anchor");
        wrapper_node.append_attribute("distT").set_value("0");
        wrapper_node.append_attribute("distB").set_value("0");
        wrapper_node.append_attribute("distL").set_value("114300");
        wrapper_node.append_attribute("distR").set_value("114300");
        wrapper_node.append_attribute("simplePos").set_value("0");
        wrapper_node.append_attribute("relativeHeight").set_value("251658240");
        
        const char* behindDoc = (position == ImagePosition::BehindText) ? "1" : "0";
        wrapper_node.append_attribute("behindDoc").set_value(behindDoc);
        wrapper_node.append_attribute("locked").set_value("0");
        wrapper_node.append_attribute("layoutInCell").set_value("1");
        wrapper_node.append_attribute("allowOverlap").set_value("1");
        
        auto simplePos = wrapper_node.append_child("wp:simplePos");
        simplePos.append_attribute("x").set_value("0");
        simplePos.append_attribute("y").set_value("0");
        
        auto positionH = wrapper_node.append_child("wp:positionH");
        positionH.append_attribute("relativeFrom").set_value("page");
        positionH.append_child("wp:posOffset").text().set(std::to_string(xOffset).c_str());
        
        auto positionV = wrapper_node.append_child("wp:positionV");
        positionV.append_attribute("relativeFrom").set_value("page");
        positionV.append_child("wp:posOffset").text().set(std::to_string(yOffset).c_str());
    }
    
    // Parse image dimensions dynamically
    auto [px_w, px_h] = getImageSize(image_path);
    long long emu_w = 4572000; // default 5 inches
    long long emu_h = 3048000; // default 3.33 inches
    if (px_w > 0 && px_h > 0) {
        emu_w = px_w * 9525LL; // assuming 96 DPI: 1 px = 914400 / 96 = 9525 EMU
        emu_h = px_h * 9525LL;
    }
    
    // Apply scale factor
    if (scale > 0 && scale != 1.0) {
        emu_w = static_cast<long long>(emu_w * scale);
        emu_h = static_cast<long long>(emu_h * scale);
    }

    auto extent = wrapper_node.append_child("wp:extent");
    extent.append_attribute("cx").set_value(std::to_string(emu_w).c_str());
    extent.append_attribute("cy").set_value(std::to_string(emu_h).c_str());
    
    auto effect = wrapper_node.append_child("wp:effectExtent");
    effect.append_attribute("l").set_value("0");
    effect.append_attribute("t").set_value("0");
    effect.append_attribute("r").set_value("0");
    effect.append_attribute("b").set_value("0");
        
    if (position != ImagePosition::Inline) {
        wrapper_node.append_child("wp:wrapNone");
    }
    
    auto docPr = wrapper_node.append_child("wp:docPr");
    docPr.append_attribute("id").set_value(std::to_string(c + 1).c_str());
    docPr.append_attribute("name").set_value(("Picture " + std::to_string(c + 1)).c_str());
    
    if (position != ImagePosition::Inline) {
        auto cNvGraphicFramePr = wrapper_node.append_child("wp:cNvGraphicFramePr");
        auto graphicFrameLocks = cNvGraphicFramePr.append_child("a:graphicFrameLocks");
        graphicFrameLocks.append_attribute("xmlns:a").set_value("http://schemas.openxmlformats.org/drawingml/2006/main");
        graphicFrameLocks.append_attribute("noChangeAspect").set_value("1");
    }
    
    auto graphic = wrapper_node.append_child("a:graphic");
    auto graphicData = graphic.append_child("a:graphicData");
    graphicData.append_attribute("uri").set_value("http://schemas.openxmlformats.org/drawingml/2006/picture");
    
    auto pic = graphicData.append_child("pic:pic");
    auto nvPicPr = pic.append_child("pic:nvPicPr");
    auto cNvPr = nvPicPr.append_child("pic:cNvPr");
    cNvPr.append_attribute("id").set_value("1"); // Use ID=1 to avoid 0 which sometimes causes issues
    cNvPr.append_attribute("name").set_value("Picture 1"); // generic name
    auto cNvPicPr = nvPicPr.append_child("pic:cNvPicPr");
    auto picLocks = cNvPicPr.append_child("a:picLocks");
    picLocks.append_attribute("noChangeAspect").set_value("1");
    picLocks.append_attribute("noChangeArrowheads").set_value("1");
    
    auto blipFill = pic.append_child("pic:blipFill");
    auto blip = blipFill.append_child("a:blip");
    blip.append_attribute("r:embed").set_value(rid.c_str());
    blip.append_attribute("cstate").set_value("print");
    auto stretch = blipFill.append_child("a:stretch");
    stretch.append_child("a:fillRect");
    
    auto spPr = pic.append_child("pic:spPr");
    auto xfrm = spPr.append_child("a:xfrm");
    auto off = xfrm.append_child("a:off");
    off.append_attribute("x").set_value("0");
    off.append_attribute("y").set_value("0");
    auto ext = xfrm.append_child("a:ext");
    ext.append_attribute("cx").set_value(std::to_string(emu_w).c_str());
    ext.append_attribute("cy").set_value(std::to_string(emu_h).c_str());
    
    auto prstGeom = spPr.append_child("a:prstGeom");
    prstGeom.append_attribute("prst").set_value("rect");
    prstGeom.append_child("a:avLst");
}

Run Paragraph::addHyperlink(gsl::czstring text, gsl::czstring url) {
    auto n = cast_node(node_);
    auto root = n.root().child("w:document");
    auto h_store = root.child("openword_hyperlinks");
    if (!h_store) h_store = root.append_child("openword_hyperlinks");
    
    auto rel_id_attr = root.attribute("openword_next_rel_id");
    if (!rel_id_attr) {
        rel_id_attr = root.append_attribute("openword_next_rel_id");
        rel_id_attr.set_value("100");
    }
    int rel_id = rel_id_attr.as_int(100);
    rel_id_attr.set_value(std::to_string(rel_id + 1).c_str());
    std::string rid_str = "rId" + std::to_string(rel_id);
    
    auto h_entry = h_store.append_child("link");
    h_entry.append_attribute("url") = url;
    h_entry.append_attribute("rId") = rid_str.c_str();
    
    auto h_node = n.append_child("w:hyperlink");
    h_node.append_attribute("r:id") = rid_str.c_str();
    h_node.append_attribute("w:history") = "1";
    
    auto r_node = h_node.append_child("w:r");
    r_node.append_child("w:rPr").append_child("w:rStyle").append_attribute("w:val") = "Hyperlink";
    
    Run run_proxy(r_node.internal_object());
    run_proxy.setText(text);
    return run_proxy;
}

Run Paragraph::addInternalLink(gsl::czstring text, gsl::czstring bookmarkName) {
    auto n = cast_node(node_);
    auto h_node = n.append_child("w:hyperlink");
    h_node.append_attribute("w:anchor") = bookmarkName;
    h_node.append_attribute("w:history") = "1";
    
    auto r_node = h_node.append_child("w:r");
    r_node.append_child("w:rPr").append_child("w:rStyle").append_attribute("w:val") = "Hyperlink";

    Run run_proxy(r_node.internal_object());
    run_proxy.setText(text);
    return run_proxy;
}

void Paragraph::insertBookmark(gsl::czstring name) {
    auto n = cast_node(node_);
    auto root = n.root().child("w:document");
    if (!root.attribute("openword_next_bkmk_id")) {
        root.append_attribute("openword_next_bkmk_id") = "0";
    }
    int bkmk_id = root.attribute("openword_next_bkmk_id").as_int(0);
    root.attribute("openword_next_bkmk_id").set_value(std::to_string(bkmk_id + 1).c_str());
    std::string id_str = std::to_string(bkmk_id);
    
    auto start = n.append_child("w:bookmarkStart");
    start.append_attribute("w:id") = id_str.c_str();
    start.append_attribute("w:name") = name;
    
    auto end = n.append_child("w:bookmarkEnd");
    end.append_attribute("w:id") = id_str.c_str();
}

void Paragraph::addFootnoteReference(int footnoteId) {
    auto n = cast_node(node_);
    auto r_node = n.append_child("w:r");
    r_node.append_child("w:rPr").append_child("w:rStyle").append_attribute("w:val") = "FootnoteReference";
    auto f_ref = r_node.append_child("w:footnoteReference");
    f_ref.append_attribute("w:id") = std::to_string(footnoteId).c_str();
}

void Paragraph::addEndnoteReference(int endnoteId) {
    auto n = cast_node(node_);
    auto r_node = n.append_child("w:r");
    r_node.append_child("w:rPr").append_child("w:rStyle").append_attribute("w:val") = "EndnoteReference";
    auto e_ref = r_node.append_child("w:endnoteReference");
    e_ref.append_attribute("w:id") = std::to_string(endnoteId).c_str();
}

int Paragraph::replaceText(const std::string& search, const std::string& replace) {
    if (search.empty()) return 0;
    
    auto p_node = cast_node(node_);
    int replace_count = 0;
    size_t search_offset = 0;
    
    while (true) {
        std::string full_text;
        struct TextRun { pugi::xml_node t_node; size_t start_idx; size_t length; std::string text; };
        std::vector<TextRun> runs;
        
        for (auto r = p_node.child("w:r"); r; r = r.next_sibling("w:r")) {
            for (auto t = r.child("w:t"); t; t = t.next_sibling("w:t")) {
                std::string val = t.text().get();
                runs.push_back({t, full_text.size(), val.size(), val});
                full_text += val;
            }
        }
        
        size_t match_pos = full_text.find(search, search_offset);
        if (match_pos == std::string::npos) break; 
        
        size_t match_end = match_pos + search.size();
        
        bool replacement_inserted = false;
        for (auto& tr : runs) {
            size_t r_start = tr.start_idx;
            size_t r_end = tr.start_idx + tr.length;
            
            // If run is completely outside the match, skip
            if (r_end <= match_pos || r_start >= match_end) continue;
            
            std::string new_val;
            
            // Preserve text before the match in this run
            if (r_start < match_pos) {
                new_val += tr.text.substr(0, match_pos - r_start);
            }
            
            // We want to insert the replacement text in the run that carries the "meat" of the text,
            // or if it spans multiple, the run that has styling. To be simple and robust:
            // We'll put it in the FIRST run that overlaps the match, BUT wait:
            // If the user did p1.addRun("{{"); p1.addRun("NAME").setBold(); p1.addRun("}}");
            // "NAME" is the middle run. If we put it in the first "{{", it loses the bold.
            // Better heuristic: Put the replacement in the run that overlaps the CENTER of the match.
            size_t match_center = match_pos + search.size() / 2;
            bool is_center_run = (r_start <= match_center && r_end > match_center);
            
            // Fallback: if we haven't inserted it by the very last overlapping run, insert it there.
            bool is_last_overlapping_run = (r_end >= match_end);

            if (!replacement_inserted && (is_center_run || is_last_overlapping_run)) {
                new_val += replace;
                replacement_inserted = true;
            }
            
            // Preserve text after the match in this run
            if (r_end > match_end) {
                new_val += tr.text.substr(match_end - r_start);
            }
            
            tr.t_node.text().set(new_val.c_str());
            
            // Enforce whitespace preservation if necessary
            if (!new_val.empty() && (new_val.front() == ' ' || new_val.back() == ' ' || new_val.find('\t') != std::string::npos || new_val.find('\n') != std::string::npos)) {
                if (!tr.t_node.attribute("xml:space")) {
                    tr.t_node.append_attribute("xml:space") = "preserve";
                }
            }
        }
        
        search_offset = match_pos + replace.size(); // Advance to prevent infinite loops
        replace_count++;
    }
    
    return replace_count;
}

void Paragraph::addEquation(const std::string& omml) {
    if (omml.empty()) return;
    auto n = cast_node(node_);
    
    pugi::xml_document math_doc;
    if (!math_doc.load_string(omml.c_str())) return;
    
    auto math_node = math_doc.child("m:oMath");
    if (!math_node) math_node = math_doc.child("m:oMathPara");
    
    if (math_node) {
        n.append_copy(math_node);
    }
}

void Paragraph::remove() {
    auto n = cast_node(node_);
    if (n && n.parent()) {
        n.parent().remove_child(n);
    }
}

Paragraph Paragraph::cloneAfter() {
    auto n = cast_node(node_);
    if (!n || !n.parent()) return Paragraph(nullptr);
    auto cloned = n.parent().insert_copy_after(n, n);
    return Paragraph(cloned.internal_object());
}

Paragraph Paragraph::insertParagraphAfter(const std::string& text) {
    auto n = cast_node(node_);
    if (!n || !n.parent()) return Paragraph(nullptr);
    auto p_node = n.parent().insert_child_after("w:p", n);
    Paragraph para(p_node.internal_object());
    if (!text.empty()) {
        para.addRun(text);
    }
    return para;
}

Table Paragraph::insertTableAfter(int rows, int cols) {
    auto n = cast_node(node_);
    if (!n || !n.parent()) return Table(nullptr);
    
    auto tbl = n.parent().insert_child_after("w:tbl", n);
    auto tblPr = tbl.append_child("w:tblPr");
    tblPr.append_child("w:tblStyle").append_attribute("w:val") = "TableGrid";
    auto tblW = tblPr.append_child("w:tblW");
    tblW.append_attribute("w:w") = "0";
    tblW.append_attribute("w:type") = "auto";
    auto tblBorders = tblPr.append_child("w:tblBorders");
    
    for (int i = 0; i < rows; ++i) {
        auto tr = tbl.append_child("w:tr");
        for (int j = 0; j < cols; ++j) {
            auto tc = tr.append_child("w:tc");
            tc.append_child("w:tcPr").append_child("w:tcW").append_attribute("w:w") = "0";
            tc.child("w:tcPr").child("w:tcW").append_attribute("w:type") = "auto";
            tc.append_child("w:p");
        }
    }
    return Table(tbl.internal_object());
}

} // namespace openword
