#include <openword/Document.h>
#include "Internal.h"

#include <pugixml.hpp>
#include <stdexcept>

namespace openword {

Cell::Cell(void* node) : node_(node) {
    Expects(node_ != nullptr);
}

Paragraph Cell::addParagraph(const std::string& text) {
    auto n = cast_node(node_);
    // Word requires a paragraph in every cell. If one exists, use it or append after
    auto p = n.child("w:p");
    if (!p) {
        p = n.append_child("w:p");
    } else if (!text.empty() && p.child("w:r")) {
        // if there's already a paragraph with text, add a new one
        p = n.append_child("w:p");
    }
    
    Paragraph para(p.internal_object());
    if (!text.empty()) {
        para.addRun(text);
    }
    return para;
}

Table::Table(void* node) : node_(node) {
    Expects(node_ != nullptr);
}

Cell Table::cell(int row, int col) {
    auto n = cast_node(node_);
    auto tr = n.child("w:tr");
    for (int i = 0; i < row && tr; ++i) {
        tr = tr.next_sibling("w:tr");
    }
    if (!tr) throw std::out_of_range("row index out of range");

    auto tc = tr.child("w:tc");
    int visual_col = 0;
    while (tc) {
        int span = 1;
        auto gridSpan = tc.child("w:tcPr").child("w:gridSpan");
        if (gridSpan) {
            span = gridSpan.attribute("w:val").as_int(1);
        }
        
        if (col >= visual_col && col < visual_col + span) {
            return Cell(tc.internal_object());
        }
        visual_col += span;
        tc = tc.next_sibling("w:tc");
    }
    throw std::out_of_range("col index out of range");
}

void Table::mergeCells(int startRow, int startCol, int endRow, int endCol) {
    Expects(startRow <= endRow);
    Expects(startCol <= endCol);
    
    auto n = cast_node(node_);
    int current_row = 0;
    
    for (auto tr = n.child("w:tr"); tr; tr = tr.next_sibling("w:tr")) {
        if (current_row >= startRow && current_row <= endRow) {
            int visual_col = 0;
            pugi::xml_node start_tc;
            std::vector<pugi::xml_node> to_delete;
            
            for (auto tc = tr.child("w:tc"); tc; tc = tc.next_sibling("w:tc")) {
                int span = 1;
                auto gs = tc.child("w:tcPr").child("w:gridSpan");
                if (gs) span = gs.attribute("w:val").as_int(1);
                
                if (visual_col == startCol) {
                    start_tc = tc;
                } else if (visual_col > startCol && visual_col <= endCol) {
                    to_delete.push_back(tc);
                }
                
                visual_col += span;
                if (visual_col > endCol) break;
            }
            
            if (start_tc) {
                auto tcPr = start_tc.child("w:tcPr");
                if (!tcPr) tcPr = start_tc.prepend_child("w:tcPr");
                
                if (startCol < endCol) {
                    int span = endCol - startCol + 1;
                    auto gridSpan = tcPr.child("w:gridSpan");
                    if (!gridSpan) gridSpan = tcPr.append_child("w:gridSpan");
                    auto attr = gridSpan.attribute("w:val");
                    if (!attr) attr = gridSpan.append_attribute("w:val");
                    attr.set_value(std::to_string(span).c_str());
                }
                
                if (startRow < endRow) {
                    auto vMerge = tcPr.child("w:vMerge");
                    if (!vMerge) vMerge = tcPr.append_child("w:vMerge");
                    
                    if (current_row == startRow) {
                        auto attr = vMerge.attribute("w:val");
                        if (!attr) attr = vMerge.append_attribute("w:val");
                        attr.set_value("restart");
                    } else {
                        auto attr = vMerge.attribute("w:val");
                        if (!attr) attr = vMerge.append_attribute("w:val");
                        attr.set_value("continue");
                        
                        // Clear contents of merged-over cells to avoid unexpected behavior
                        while (start_tc.child("w:p")) start_tc.remove_child("w:p");
                        auto p = start_tc.append_child("w:p");
                        p.append_child("w:pPr"); 
                    }
                }
                
                // Delete horizontally merged cells
                for (auto& del_node : to_delete) {
                    tr.remove_child(del_node);
                }
            }
        }
        current_row++;
    }
}

} // namespace openword