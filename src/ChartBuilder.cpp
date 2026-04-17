#include "ChartBuilder.h"
#include <sstream>

namespace openword {
namespace internal {

ChartBuilder::ChartBuilder() {
    initializeBaseDom();
}

void ChartBuilder::initializeBaseDom() {
    // Add XML declaration
    pugi::xml_node decl = doc_.append_child(pugi::node_declaration);
    decl.append_attribute("version") = "1.0";
    decl.append_attribute("encoding") = "UTF-8";
    decl.append_attribute("standalone") = "yes";

    // Initialize root node chartSpace
    chartSpace_ = doc_.append_child("c:chartSpace");
    chartSpace_.append_attribute("xmlns:c") = openxml_ns::c;
    chartSpace_.append_attribute("xmlns:a") = openxml_ns::a;
    chartSpace_.append_attribute("xmlns:r") = openxml_ns::r;

    chartSpace_.append_child("c:date1904").append_attribute("val") = "0";
    chartSpace_.append_child("c:lang").append_attribute("val") = "en-US";
    chartSpace_.append_child("c:roundedCorners").append_attribute("val") = "0";

    pugi::xml_node altContent = chartSpace_.append_child("mc:AlternateContent");
    altContent.append_attribute("xmlns:mc") = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    pugi::xml_node choice = altContent.append_child("mc:Choice");
    choice.append_attribute("Requires") = "c14";
    choice.append_attribute("xmlns:c14") = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart";
    choice.append_child("c14:style").append_attribute("val") = "102";
    pugi::xml_node fallback = altContent.append_child("mc:Fallback");
    fallback.append_child("c:style").append_attribute("val") = "2";

    chart_ = chartSpace_.append_child("c:chart");
}

ChartBuilder& ChartBuilder::setType(ChartType type) {
    type_ = type;
    
    // plotArea is added here if not exists
    if (!plotArea_) {
        plotArea_ = chart_.append_child("c:plotArea");
        plotArea_.append_child("c:layout");
    }

    std::string type_tag;
    if (type == ChartType::Bar)
        type_tag = "c:barChart";
    else if (type == ChartType::Line)
        type_tag = "c:lineChart";
    else if (type == ChartType::Pie)
        type_tag = "c:pieChart";

    typeGroup_ = plotArea_.append_child(type_tag.c_str());

    if (type == ChartType::Bar) {
        typeGroup_.append_child("c:barDir").append_attribute("val") = "col";
        typeGroup_.append_child("c:grouping").append_attribute("val") = "clustered";
    }

    return *this;
}

ChartBuilder& ChartBuilder::setTitle(const std::string& title) {
    if (title.empty()) return *this;

    // c:title
    pugi::xml_node titleNode = chart_.insert_child_before("c:title", plotArea_);
    
    pugi::xml_node tx = titleNode.append_child("c:tx");
    pugi::xml_node rich = tx.append_child("c:rich");
    rich.append_child("a:bodyPr");
    rich.append_child("a:lstStyle");
    
    pugi::xml_node p = rich.append_child("a:p");
    p.append_child("a:pPr").append_child("a:defRPr");
    
    pugi::xml_node r = p.append_child("a:r");
    r.append_child("a:rPr").append_attribute("lang") = "en-US";
    r.append_child("a:t").text().set(title.c_str());

    titleNode.append_child("c:layout");
    return *this;
}

ChartBuilder& ChartBuilder::setLegend(LegendPosition pos) {
    if (pos == LegendPosition::None) return *this;

    pugi::xml_node legend = chart_.append_child("c:legend");
    pugi::xml_node legendPos = legend.append_child("c:legendPos");
    
    std::string pos_val = "b";
    if (pos == LegendPosition::Top) pos_val = "t";
    else if (pos == LegendPosition::Left) pos_val = "l";
    else if (pos == LegendPosition::Right) pos_val = "r";

    legendPos.append_attribute("val") = pos_val.c_str();
    legend.append_child("c:overlay").append_attribute("val") = "0";

    return *this;
}

ChartBuilder& ChartBuilder::setPlotVisOnly(bool val) {
    chart_.append_child("c:plotVisOnly").append_attribute("val") = val ? "1" : "0";
    chart_.append_child("c:dispBlanksAs").append_attribute("val") = "gap";
    chart_.append_child("c:showDLblsOverMax").append_attribute("val") = "0";
    return *this;
}

ChartBuilder& ChartBuilder::setOptions(const ChartOptions& options) {
    options_ = options;
    return *this;
}

ChartBuilder& ChartBuilder::addSeries(const ChartSeries& series) {
    if (!typeGroup_) {
        // Fallback or error, type not set
        return *this;
    }

    pugi::xml_node ser = typeGroup_.append_child("c:ser");
    ser.append_child("c:idx").append_attribute("val") = seriesCount_;
    ser.append_child("c:order").append_attribute("val") = seriesCount_;

    pugi::xml_node tx = ser.append_child("c:tx");
    tx.append_child("c:v").text().set(series.name.c_str());

    // Color
    if (!series.colorHex.empty()) {
        pugi::xml_node spPr = ser.append_child("c:spPr");
        pugi::xml_node solidFill = spPr.append_child("a:solidFill");
        solidFill.append_child("a:srgbClr").append_attribute("val") = series.colorHex.c_str();
    }

    // Data Labels
    if (options_.showDataLabels) {
        pugi::xml_node dLbls = ser.append_child("c:dLbls");
        dLbls.append_child("c:showLegendKey").append_attribute("val") = "0";
        dLbls.append_child("c:showVal").append_attribute("val") = "1";
        dLbls.append_child("c:showCatName").append_attribute("val") = "0";
        dLbls.append_child("c:showSerName").append_attribute("val") = "0";
        dLbls.append_child("c:showPercent").append_attribute("val") = "0";
        dLbls.append_child("c:showBubbleSize").append_attribute("val") = "0";
    }

    // Categories
    pugi::xml_node cat = ser.append_child("c:cat");
    pugi::xml_node strLit = cat.append_child("c:strLit");
    buildStringCache(strLit, series.categories);

    // Values
    pugi::xml_node val = ser.append_child("c:val");
    pugi::xml_node numLit = val.append_child("c:numLit");
    numLit.append_child("c:formatCode").text().set("General");
    buildNumberCache(numLit, series.values);

    seriesCount_++;
    return *this;
}

pugi::xml_node ChartBuilder::buildStringCache(pugi::xml_node parent, const std::vector<std::string>& data) {
    parent.append_child("c:ptCount").append_attribute("val") = std::to_string(data.size()).c_str();
    for (size_t i = 0; i < data.size(); ++i) {
        pugi::xml_node pt = parent.append_child("c:pt");
        pt.append_attribute("idx") = std::to_string(i).c_str();
        pt.append_child("c:v").text().set(data[i].c_str());
    }
    return parent;
}

pugi::xml_node ChartBuilder::buildNumberCache(pugi::xml_node parent, const std::vector<double>& data) {
    parent.append_child("c:ptCount").append_attribute("val") = std::to_string(data.size()).c_str();
    for (size_t i = 0; i < data.size(); ++i) {
        pugi::xml_node pt = parent.append_child("c:pt");
        pt.append_attribute("idx") = std::to_string(i).c_str();
        
        // Convert double to string safely
        std::stringstream ss;
        ss << data[i];
        pt.append_child("c:v").text().set(ss.str().c_str());
    }
    return parent;
}

void ChartBuilder::buildAxis() {
    if (!plotArea_ || !typeGroup_) return;

    if (type_ == ChartType::Bar || type_ == ChartType::Line) {
        typeGroup_.append_child("c:axId").append_attribute("val") = "1936577344";
        typeGroup_.append_child("c:axId").append_attribute("val") = "1959785600";

        // catAx
        pugi::xml_node catAx = plotArea_.append_child("c:catAx");
        catAx.append_child("c:axId").append_attribute("val") = "1936577344";
        catAx.append_child("c:scaling").append_child("c:orientation").append_attribute("val") = "minMax";
        catAx.append_child("c:delete").append_attribute("val") = "0";
        catAx.append_child("c:axPos").append_attribute("val") = "b";
        catAx.append_child("c:majorTickMark").append_attribute("val") = "none";
        catAx.append_child("c:minorTickMark").append_attribute("val") = "none";
        catAx.append_child("c:tickLblPos").append_attribute("val") = "nextTo";
        catAx.append_child("c:crossAx").append_attribute("val") = "1959785600";
        catAx.append_child("c:crosses").append_attribute("val") = "autoZero";
        catAx.append_child("c:auto").append_attribute("val") = "1";
        catAx.append_child("c:lblAlgn").append_attribute("val") = "ctr";
        catAx.append_child("c:lblOffset").append_attribute("val") = "100";
        catAx.append_child("c:noMultiLvlLbl").append_attribute("val") = "0";

        // valAx
        pugi::xml_node valAx = plotArea_.append_child("c:valAx");
        valAx.append_child("c:axId").append_attribute("val") = "1959785600";
        valAx.append_child("c:scaling").append_child("c:orientation").append_attribute("val") = "minMax";
        valAx.append_child("c:delete").append_attribute("val") = "0";
        valAx.append_child("c:axPos").append_attribute("val") = "l";
        valAx.append_child("c:majorGridlines");
        valAx.append_child("c:majorTickMark").append_attribute("val") = "none";
        valAx.append_child("c:minorTickMark").append_attribute("val") = "none";
        valAx.append_child("c:tickLblPos").append_attribute("val") = "nextTo";
        valAx.append_child("c:crossAx").append_attribute("val") = "1936577344";
        valAx.append_child("c:crosses").append_attribute("val") = "autoZero";
        valAx.append_child("c:crossBetween").append_attribute("val") = "between";
    }
}

std::string ChartBuilder::buildXml() {
    buildAxis(); // Add axes right before building
    setLegend(options_.legendPos); // Add legend
    setPlotVisOnly(true); // Final plot setup

    std::stringstream ss;
    // Write out the XML. pugi::format_raw will output exactly what we've constructed.
    doc_.save(ss, "  ", pugi::format_default | pugi::format_save_file_text, pugi::encoding_utf8);
    return ss.str();
}

} // namespace internal
} // namespace openword