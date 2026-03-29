#include <openword/Document.h>
#include "Internal.h"

#include <pugixml.hpp>
#include <stdexcept>
#include <string>

namespace openword {

static const char* borderStyleToString(BorderStyle style) {
    switch (style) {
        case BorderStyle::None: return "none";
        case BorderStyle::Single: return "single";
        case BorderStyle::Dashed: return "dashed";
        case BorderStyle::Dotted: return "dotted";
        case BorderStyle::Double: return "double";
        case BorderStyle::Thick: return "thick";
        default: return "single";
    }
}

static void applyBorder(pugi::xml_node borders, const char* borderName, const BorderSettings& bs) {
    auto b = borders.child(borderName);
    if (!b) b = borders.append_child(borderName);
    b.remove_attribute("w:val"); b.append_attribute("w:val") = borderStyleToString(bs.style);
    b.remove_attribute("w:sz"); b.append_attribute("w:sz") = std::to_string(bs.size).c_str();
    b.remove_attribute("w:space"); b.append_attribute("w:space") = "0";
    b.remove_attribute("w:color"); b.append_attribute("w:color") = bs.color.c_str();
}

Row::Row(void* node) : node_(node) {
    Expects(node_ != nullptr);
}

Row& Row::setHeight(int twips, HeightRule rule) {
    auto n = cast_node(node_);
    auto trPr = n.child("w:trPr");
    if (!trPr) trPr = n.prepend_child("w:trPr");
    auto trHeight = trPr.child("w:trHeight");
    if (!trHeight) trHeight = trPr.append_child("w:trHeight");
    
    trHeight.remove_attribute("w:val");
    trHeight.append_attribute("w:val") = std::to_string(twips).c_str();
    
    trHeight.remove_attribute("w:hRule");
    const char* ruleStr = "atLeast";
    if (rule == HeightRule::Exact) ruleStr = "exact";
    else if (rule == HeightRule::Auto) ruleStr = "auto";
    trHeight.append_attribute("w:hRule") = ruleStr;
    
    return *this;
}

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

Cell& Cell::setVertAlign(VerticalAlignment align) {
    auto n = cast_node(node_);
    auto tcPr = n.child("w:tcPr");
    if (!tcPr) tcPr = n.prepend_child("w:tcPr");
    
    auto vAlign = tcPr.child("w:vAlign");
    if (!vAlign) vAlign = tcPr.append_child("w:vAlign");
    
    vAlign.remove_attribute("w:val");
    const char* val = "center";
    if (align == VerticalAlignment::Top) val = "top";
    else if (align == VerticalAlignment::Bottom) val = "bottom";
    vAlign.append_attribute("w:val") = val;
    
    return *this;
}

Cell& Cell::setShading(const std::string& hexColor) {
    auto n = cast_node(node_);
    auto tcPr = n.child("w:tcPr");
    if (!tcPr) tcPr = n.prepend_child("w:tcPr");
    
    auto shd = tcPr.child("w:shd");
    if (!shd) shd = tcPr.append_child("w:shd");
    
    shd.remove_attribute("w:val"); shd.append_attribute("w:val") = "clear";
    shd.remove_attribute("w:color"); shd.append_attribute("w:color") = "auto";
    shd.remove_attribute("w:fill"); shd.append_attribute("w:fill") = hexColor.c_str();
    
    return *this;
}

Cell& Cell::setWidth(int twips, const std::string& type) {
    auto n = cast_node(node_);
    auto tcPr = n.child("w:tcPr");
    if (!tcPr) tcPr = n.prepend_child("w:tcPr");
    
    auto tcW = tcPr.child("w:tcW");
    if (!tcW) tcW = tcPr.append_child("w:tcW");
    
    tcW.remove_attribute("w:w"); tcW.append_attribute("w:w") = std::to_string(twips).c_str();
    tcW.remove_attribute("w:type"); tcW.append_attribute("w:type") = type.c_str();
    
    return *this;
}

Cell& Cell::setBorders(const BorderSettings& top, const BorderSettings& bottom, const BorderSettings& left, const BorderSettings& right) {
    auto n = cast_node(node_);
    auto tcPr = n.child("w:tcPr");
    if (!tcPr) tcPr = n.prepend_child("w:tcPr");
    
    auto tcBorders = tcPr.child("w:tcBorders");
    if (!tcBorders) tcBorders = tcPr.append_child("w:tcBorders");
    
    applyBorder(tcBorders, "w:top", top);
    applyBorder(tcBorders, "w:bottom", bottom);
    applyBorder(tcBorders, "w:left", left);
    applyBorder(tcBorders, "w:right", right);
    
    return *this;
}

Cell& Cell::setBorders(const BorderSettings& all) {
    return setBorders(all, all, all, all);
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

Row Table::row(int rowIndex) {
    auto n = cast_node(node_);
    auto tr = n.child("w:tr");
    for (int i = 0; i < rowIndex && tr; ++i) {
        tr = tr.next_sibling("w:tr");
    }
    if (!tr) throw std::out_of_range("row index out of range");
    return Row(tr.internal_object());
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

Table& Table::setBorders(const BorderSettings& top, const BorderSettings& bottom, const BorderSettings& left, const BorderSettings& right, const BorderSettings& insideH, const BorderSettings& insideV) {
    auto n = cast_node(node_);
    auto tblPr = n.child("w:tblPr");
    if (!tblPr) tblPr = n.prepend_child("w:tblPr");
    
    auto tblBorders = tblPr.child("w:tblBorders");
    if (!tblBorders) tblBorders = tblPr.append_child("w:tblBorders");
    
    applyBorder(tblBorders, "w:top", top);
    applyBorder(tblBorders, "w:bottom", bottom);
    applyBorder(tblBorders, "w:left", left);
    applyBorder(tblBorders, "w:right", right);
    applyBorder(tblBorders, "w:insideH", insideH);
    applyBorder(tblBorders, "w:insideV", insideV);
    
    return *this;
}

Table& Table::setBorders(const BorderSettings& outer, const BorderSettings& inner) {
    return setBorders(outer, outer, outer, outer, inner, inner);
}

Table& Table::setBorders(const BorderSettings& all) {
    return setBorders(all, all, all, all, all, all);
}

Table& Table::setColumnWidth(int colIndex, int twips) {
    auto n = cast_node(node_);
    auto tblGrid = n.child("w:tblGrid");
    if (!tblGrid) tblGrid = n.append_child("w:tblGrid");
    
    auto gridCol = tblGrid.child("w:gridCol");
    for (int i = 0; i < colIndex && gridCol; ++i) {
        gridCol = gridCol.next_sibling("w:gridCol");
    }
    if (!gridCol) {
        int currentCols = 0;
        for (auto c : tblGrid.children("w:gridCol")) currentCols++;
        for (int i = currentCols; i <= colIndex; ++i) {
            gridCol = tblGrid.append_child("w:gridCol");
        }
    }
    
    gridCol.remove_attribute("w:w");
    gridCol.append_attribute("w:w") = std::to_string(twips).c_str();
    
    return *this;
}

Table& Table::setColumnWidths(const std::vector<int>& twipsList) {
    for (size_t i = 0; i < twipsList.size(); ++i) {
        setColumnWidth(static_cast<int>(i), twipsList[i]);
    }
    return *this;
}

Table& Table::setAlignment(gsl::czstring align) {
    auto n = cast_node(node_);
    auto tblPr = n.child("w:tblPr");
    if (!tblPr) tblPr = n.prepend_child("w:tblPr");
    auto jc = tblPr.child("w:jc");
    if (!jc) jc = tblPr.append_child("w:jc");
    jc.remove_attribute("w:val");
    jc.append_attribute("w:val") = align;
    return *this;
}

int Cell::replaceText(const std::string& search, const std::string& replace) {
    int count = 0;
    auto n = cast_node(node_);
    for (auto child : n.children()) {
        std::string name = child.name();
        if (name == "w:p") {
            count += Paragraph(child.internal_object()).replaceText(search, replace);
        } else if (name == "w:tbl") {
            count += Table(child.internal_object()).replaceText(search, replace);
        }
    }
    return count;
}

int Table::replaceText(const std::string& search, const std::string& replace) {
    int count = 0;
    auto n = cast_node(node_);
    for (auto tr = n.child("w:tr"); tr; tr = tr.next_sibling("w:tr")) {
        for (auto tc = tr.child("w:tc"); tc; tc = tc.next_sibling("w:tc")) {
            count += Cell(tc.internal_object()).replaceText(search, replace);
        }
    }
    return count;
}


int Row::replaceText(const std::string& search, const std::string& replace) {
    int count = 0;
    for (auto& c : cells()) count += c.replaceText(search, replace);
    return count;
}

// --- Row Extraction ---

std::vector<Cell> Row::cells() const {
    std::vector<Cell> result;
    auto n = cast_node(node_);
    for (auto tc : n.children("w:tc")) {
        result.push_back(Cell(tc.internal_object()));
    }
    return result;
}

std::string Row::text() const {
    std::string result;
    for (auto& c : cells()) {
        result += c.text() + " ";
    }
    if (!result.empty()) result.pop_back(); // Remove trailing space
    return result;
}

// --- Cell Extraction ---

std::vector<Paragraph> Cell::paragraphs() const {
    std::vector<Paragraph> result;
    auto n = cast_node(node_);
    for (auto p : n.children("w:p")) {
        result.push_back(Paragraph(p.internal_object()));
    }
    return result;
}

std::string Cell::text() const {
    std::string result;
    for (auto& p : paragraphs()) {
        result += p.text() + "\n";
    }
    if (!result.empty()) result.pop_back(); // Remove trailing newline
    return result;
}

// --- Table Extraction ---

std::vector<Row> Table::rows() const {
    std::vector<Row> result;
    auto n = cast_node(node_);
    for (auto tr : n.children("w:tr")) {
        result.push_back(Row(tr.internal_object()));
    }
    return result;
}

std::string Table::text() const {
    std::string result;
    for (auto& r : rows()) {
        result += r.text() + "\n";
    }
    if (!result.empty()) result.pop_back();
    return result;
}


void Row::remove() {
    auto n = cast_node(node_);
    if (n && n.parent()) {
        n.parent().remove_child(n);
    }
}

Row Row::cloneAfter() {
    auto n = cast_node(node_);
    if (!n || !n.parent()) return Row(nullptr);
    auto cloned = n.parent().insert_copy_after(n, n);
    return Row(cloned.internal_object());
}

void Table::remove() {
    auto n = cast_node(node_);
    if (n && n.parent()) {
        n.parent().remove_child(n);
    }
}

Table Table::cloneAfter() {
    auto n = cast_node(node_);
    if (!n || !n.parent()) return Table(nullptr);
    auto cloned = n.parent().insert_copy_after(n, n);
    return Table(cloned.internal_object());
}

Paragraph Table::insertParagraphAfter(const std::string& text) {
    auto n = cast_node(node_);
    if (!n || !n.parent()) return Paragraph(nullptr);
    auto p_node = n.parent().insert_child_after("w:p", n);
    Paragraph para(p_node.internal_object());
    if (!text.empty()) {
        para.addRun(text);
    }
    return para;
}

} // namespace openword

