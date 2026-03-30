#include "openword/Document.h"
#include "Internal.h"
#include <string>

namespace openword {

static const char* numberingFormatToString(NumberingFormat fmt) {
    switch (fmt) {
        case NumberingFormat::Decimal: return "decimal";
        case NumberingFormat::LowerLetter: return "lowerLetter";
        case NumberingFormat::UpperLetter: return "upperLetter";
        case NumberingFormat::LowerRoman: return "lowerRoman";
        case NumberingFormat::UpperRoman: return "upperRoman";
        case NumberingFormat::Bullet: return "bullet";
        default: return "decimal";
    }
}

AbstractNumbering::AbstractNumbering(void* node) : node_(node) {}

AbstractNumbering& AbstractNumbering::addLevel(const ListLevel& level) {
    auto n = cast_node(node_);
    auto lvl = n.append_child("w:lvl");
    lvl.append_attribute("w:ilvl") = std::to_string(level.levelIndex).c_str();
    
    lvl.append_child("w:start").append_attribute("w:val") = "1";
    lvl.append_child("w:numFmt").append_attribute("w:val") = numberingFormatToString(level.format);
    lvl.append_child("w:lvlText").append_attribute("w:val") = level.text.c_str();
    lvl.append_child("w:lvlJc").append_attribute("w:val") = level.jc.c_str();
    
    auto pPr = lvl.append_child("w:pPr");
    auto ind = pPr.append_child("w:ind");
    ind.append_attribute("w:left") = std::to_string(level.indentTwips).c_str();
    ind.append_attribute("w:hanging") = std::to_string(level.hangingTwips).c_str();
    
    if (!level.fontAscii.empty()) {
        auto rPr = lvl.append_child("w:rPr");
        auto rFonts = rPr.append_child("w:rFonts");
        rFonts.append_attribute("w:ascii") = level.fontAscii.c_str();
        rFonts.append_attribute("w:hAnsi") = level.fontAscii.c_str();
        rFonts.append_attribute("w:hint") = "default";
    }
    
    return *this;
}

NumberingCollection::NumberingCollection(void* node) : node_(node) {}

AbstractNumbering NumberingCollection::addAbstractNumbering(int abstractNumId) {
    auto n = cast_node(node_);
    
    // Abstract nums must appear BEFORE <w:num> definitions in numbering.xml
    auto absNum = n.append_child("w:abstractNum");
    absNum.append_attribute("w:abstractNumId") = std::to_string(abstractNumId).c_str();
    absNum.append_child("w:multiLevelType").append_attribute("w:val") = "multilevel";
    
    // Sort logic to ensure <w:abstractNum> is before <w:num>
    auto firstNum = n.child("w:num");
    if (firstNum) {
        n.insert_move_before(absNum, firstNum);
    }
    
    return AbstractNumbering(absNum.internal_object());
}

int NumberingCollection::addList(int abstractNumId, int restartNumId) {
    auto n = cast_node(node_);
    
    int newNumId = 1;
    for (auto num = n.child("w:num"); num; num = num.next_sibling("w:num")) {
        int id = num.attribute("w:numId").as_int(0);
        if (id >= newNumId) newNumId = id + 1;
    }
    
    auto num = n.append_child("w:num");
    num.append_attribute("w:numId") = std::to_string(newNumId).c_str();
    num.append_child("w:abstractNumId").append_attribute("w:val") = std::to_string(abstractNumId).c_str();
    
    if (restartNumId > 0) {
        auto lvlOverride = num.append_child("w:lvlOverride");
        lvlOverride.append_attribute("w:ilvl") = "0";
        lvlOverride.append_child("w:startOverride").append_attribute("w:val") = "1";
    }
    
    return newNumId;
}

} // namespace openword


int openword::NumberingCollection::addBulletList() {
    auto n = cast_node(node_);
    
    // Check if we already created the default bullet abstract num (we use id 9001)
    int abstractId = 9001;
    bool found = false;
    for (auto absNum : n.children("w:abstractNum")) {
        if (absNum.attribute("w:abstractNumId").as_int() == abstractId) {
            found = true;
            break;
        }
    }
    
    if (!found) {
        auto abs = addAbstractNumbering(abstractId);
        
        ListLevel lvl0;
        lvl0.levelIndex = 0;
        lvl0.format = NumberingFormat::Bullet;
        lvl0.text = ""; // UTF-8 bullet
        lvl0.fontAscii = "Symbol";
        lvl0.indentTwips = 720;
        lvl0.hangingTwips = 360;
        abs.addLevel(lvl0);
        
        ListLevel lvl1;
        lvl1.levelIndex = 1;
        lvl1.format = NumberingFormat::Bullet;
        lvl1.text = "o"; // Circle
        lvl1.fontAscii = "Courier New";
        lvl1.indentTwips = 1440;
        lvl1.hangingTwips = 360;
        abs.addLevel(lvl1);
        
        ListLevel lvl2;
        lvl2.levelIndex = 2;
        lvl2.format = NumberingFormat::Bullet;
        lvl2.text = ""; // Square
        lvl2.fontAscii = "Wingdings";
        lvl2.indentTwips = 2160;
        lvl2.hangingTwips = 360;
        abs.addLevel(lvl2);
    }
    
    return addList(abstractId);
}

int openword::NumberingCollection::addNumberedList() {
    auto n = cast_node(node_);
    
    // Check if we already created the default numbered abstract num (we use id 9002)
    int abstractId = 9002;
    bool found = false;
    for (auto absNum : n.children("w:abstractNum")) {
        if (absNum.attribute("w:abstractNumId").as_int() == abstractId) {
            found = true;
            break;
        }
    }
    
    if (!found) {
        auto abs = addAbstractNumbering(abstractId);
        
        ListLevel lvl0;
        lvl0.levelIndex = 0;
        lvl0.format = NumberingFormat::Decimal;
        lvl0.text = "%1.";
        lvl0.indentTwips = 720;
        lvl0.hangingTwips = 360;
        abs.addLevel(lvl0);
        
        ListLevel lvl1;
        lvl1.levelIndex = 1;
        lvl1.format = NumberingFormat::LowerLetter;
        lvl1.text = "%2.";
        lvl1.indentTwips = 1440;
        lvl1.hangingTwips = 360;
        abs.addLevel(lvl1);
        
        ListLevel lvl2;
        lvl2.levelIndex = 2;
        lvl2.format = NumberingFormat::LowerRoman;
        lvl2.text = "%3.";
        lvl2.indentTwips = 2160;
        lvl2.hangingTwips = 360;
        abs.addLevel(lvl2);
    }
    
    return addList(abstractId);
}
