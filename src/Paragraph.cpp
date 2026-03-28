#include <openword/Document.h>
#include "Internal.h"

#include <pugixml.hpp>
#include <fstream>

namespace openword {

// Returns {width, height} in pixels
static std::pair<int, int> getImageSize(const std::string& path) {
    std::ifstream file(path, std::ios::binary);
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

void Paragraph::addImage(gsl::czstring image_path, double scale) {
    auto n = cast_node(node_);
    auto root = n.root().child("w:document");
    auto media_store = root.child("openword_media");
    if (!media_store) {
        media_store = root.append_child("openword_media");
    }

    int c = 0;
    for (auto child : media_store.children()) c++;
    std::string rid = "rIdMedia" + std::to_string(c + 1);

    auto media_entry = media_store.append_child("media");
    media_entry.append_attribute("path").set_value(image_path);
    media_entry.append_attribute("rId").set_value(rid.c_str());

    auto r = n.append_child("w:r");
    auto drawing = r.append_child("w:drawing");
    auto inline_node = drawing.append_child("wp:inline");
    
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

    auto extent = inline_node.append_child("wp:extent");
    extent.append_attribute("cx").set_value(std::to_string(emu_w).c_str());
    extent.append_attribute("cy").set_value(std::to_string(emu_h).c_str());

    auto docPr = inline_node.append_child("wp:docPr");
    docPr.append_attribute("id").set_value("1");
    docPr.append_attribute("name").set_value("Picture 1");
    
    auto cNvGraphicFramePr = inline_node.append_child("wp:cNvGraphicFramePr");
    auto graphicFrameLocks = cNvGraphicFramePr.append_child("a:graphicFrameLocks");
    graphicFrameLocks.append_attribute("xmlns:a").set_value("http://schemas.openxmlformats.org/drawingml/2006/main");
    graphicFrameLocks.append_attribute("noChangeAspect").set_value("1");

    auto graphic = inline_node.append_child("a:graphic");
    graphic.append_attribute("xmlns:a").set_value("http://schemas.openxmlformats.org/drawingml/2006/main");
    
    auto graphicData = graphic.append_child("a:graphicData");
    graphicData.append_attribute("uri").set_value("http://schemas.openxmlformats.org/drawingml/2006/picture");
    
    auto pic = graphicData.append_child("pic:pic");
    pic.append_attribute("xmlns:pic").set_value("http://schemas.openxmlformats.org/drawingml/2006/picture");
    auto nvPicPr = pic.append_child("pic:nvPicPr");
    auto cNvPr = nvPicPr.append_child("pic:cNvPr");
    cNvPr.append_attribute("id").set_value("0");
    cNvPr.append_attribute("name").set_value("Image");
    nvPicPr.append_child("pic:cNvPicPr");

    auto blipFill = pic.append_child("pic:blipFill");
    auto blip = blipFill.append_child("a:blip");
    blip.append_attribute("r:embed").set_value(rid.c_str());
    blipFill.append_child("a:stretch").append_child("a:fillRect");

    auto spPr = pic.append_child("pic:spPr");
    auto xfrm = spPr.append_child("a:xfrm");
    auto off = xfrm.append_child("a:off");
    off.append_attribute("x").set_value("0");
    off.append_attribute("y").set_value("0");
    auto ext2 = xfrm.append_child("a:ext");
    ext2.append_attribute("cx").set_value(std::to_string(emu_w).c_str());
    ext2.append_attribute("cy").set_value(std::to_string(emu_h).c_str());
    spPr.append_child("a:prstGeom").append_attribute("prst").set_value("rect");
}

} // namespace openword