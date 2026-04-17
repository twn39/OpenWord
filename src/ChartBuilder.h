#pragma once

#include "openword/Document.h"
#include <gsl/span>
#include <pugixml.hpp>
#include <string>
#include <vector>

namespace openword {
namespace internal {

namespace openxml_ns {
    constexpr gsl::czstring c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    constexpr gsl::czstring a = "http://schemas.openxmlformats.org/drawingml/2006/main";
    constexpr gsl::czstring r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
}

class ChartBuilder {
public:
    ChartBuilder();

    ChartBuilder& setType(ChartType type);
    ChartBuilder& setTitle(const std::string& title);
    ChartBuilder& setLegend(LegendPosition pos);
    ChartBuilder& setPlotVisOnly(bool val = true);
    ChartBuilder& setOptions(const ChartOptions& options);

    ChartBuilder& addSeries(const ChartSeries& series);

    std::string buildXml();

private:
    pugi::xml_document doc_;
    pugi::xml_node chartSpace_;
    pugi::xml_node chart_;
    pugi::xml_node plotArea_;
    pugi::xml_node typeGroup_; // e.g., <c:barChart>, <c:lineChart>

    ChartType type_{ChartType::Bar};
    int seriesCount_{0};
    ChartOptions options_;

    void initializeBaseDom();
    void buildAxis();
    pugi::xml_node buildStringCache(pugi::xml_node parent, gsl::span<const std::string> data);
    pugi::xml_node buildNumberCache(pugi::xml_node parent, gsl::span<const double> data);
};

} // namespace internal
} // namespace openword
