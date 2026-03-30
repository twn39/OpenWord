#pragma once

#include <filesystem>
#include <fstream>
#include <gsl/gsl>
#include <map>
#include <pugixml.hpp>
#include <string>

namespace openword {

// Helper to open std::ifstream safely using a UTF-8 path across platforms (especially Windows)
inline void open_ifstream(std::ifstream &stream, const std::string &utf8_path,
                          std::ios_base::openmode mode = std::ios_base::in) {
    std::filesystem::path p = std::filesystem::u8path(utf8_path);
    stream.open(p, mode);
}

inline pugi::xml_node cast_node(void *ptr) {
    Expects(ptr != nullptr);
    return pugi::xml_node(static_cast<pugi::xml_node_struct *>(ptr));
}

class RelationshipManager {
  public:
    std::string addRelationship(const std::string &type, const std::string &target, const std::string &forceId = "",
                                const std::string &targetMode = "") {
        std::string rId = forceId.empty() ? "rId" + std::to_string(nextId_++) : forceId;
        relationships_[rId] = {type, target, targetMode};
        return rId;
    }
    bool empty() const { return relationships_.empty(); }
    void serialize(pugi::xml_node parent) const {
        for (const auto &[id, data] : relationships_) {
            auto rel = parent.append_child("Relationship");
            rel.append_attribute("Id") = id.c_str();
            rel.append_attribute("Type") = data.type.c_str();
            rel.append_attribute("Target") = data.target.c_str();
            if (!data.targetMode.empty()) {
                rel.append_attribute("TargetMode") = data.targetMode.c_str();
            }
        }
    }
    void setNextId(int id) { nextId_ = id; }

  private:
    struct RelData {
        std::string type;
        std::string target;
        std::string targetMode;
    };
    int nextId_ = 1;
    std::map<std::string, RelData> relationships_;
};

// Math conversion helper
std::string convert_latex_to_omml(const std::string &latex);
std::string convert_mathml_to_omml(const std::string &mathml);

} // namespace openword
