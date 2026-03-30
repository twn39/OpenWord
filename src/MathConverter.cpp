#include "Internal.h"
#include <algorithm>
#include <fmt/core.h>
#include <gsl/gsl>

// Generated embedded resource
#include "MML2OMML_XSL.h"

#include <libxml/parser.h>
#include <libxml/tree.h>
#include <libxslt/transform.h>
#include <libxslt/xsltutils.h>
#include <pugixml.hpp>
#include <set>
#include <string>
#include <vector>

// C-linkage to Rust functions
extern "C" {
char *latex_to_mathml_c(const char *latex);
void free_rust_string(char *s);
}

namespace openword {

namespace {

struct xml_string_writer : pugi::xml_writer {
    std::string result;
    void write(const void *data, size_t size) override { result.append(static_cast<const char *>(data), size); }
};

/**
 * @brief Cleans up MathML to ensure high compatibility with XSLT.
 */
std::string sanitize_mathml(const std::string &input) {
    pugi::xml_document doc;
    if (!doc.load_string(input.c_str())) {
        return input;
    }

    // 1. Remove all mspace and layout hints that often break conversion
    std::array<const char *, 5> unwanted = {"mspace", "maction", "merror", "mphantom", "mpath"};
    for (const char *name : unwanted) {
        auto matches = doc.select_nodes((std::string("//*[local-name()='") + name + "']").c_str());
        for (auto m : matches)
            m.node().parent().remove_child(m.node());
    }

    // 2. Ensure mml: prefix for all elements and fix tex2math dot mappings
    struct prefixer : pugi::xml_tree_walker {
        bool for_each(pugi::xml_node &node) override {
            if (node.type() == pugi::node_element) {
                std::string name = node.name();
                
                // Fix tex2math dots: convert <mi>\dots</mi> to <mo>&#x22ef;</mo> (⋯) etc.
                if (name == "mi" || name == "mml:mi") {
                    std::string val = node.child_value();
                    if (val == "\\dots" || val == "\\cdots" || val == "⋯") {
                        node.set_name(name == "mi" ? "mo" : "mml:mo");
                        node.first_child().set_value("⋯");
                    } else if (val == "\\vdots" || val == "\\varvdots" || val == "⋮") {
                        node.set_name(name == "mi" ? "mo" : "mml:mo");
                        node.first_child().set_value("⋮");
                    } else if (val == "\\ddots" || val == "⋱") {
                        node.set_name(name == "mi" ? "mo" : "mml:mo");
                        node.first_child().set_value("⋱");
                    } else if (val == "\\implies") {
                        node.set_name(name == "mi" ? "mo" : "mml:mo");
                        node.first_child().set_value("⟹");
                    } else if (val == "\\impliedby") {
                        node.set_name(name == "mi" ? "mo" : "mml:mo");
                        node.first_child().set_value("⟸");
                    } else if (val == "\\iff") {
                        node.set_name(name == "mi" ? "mo" : "mml:mo");
                        node.first_child().set_value("⟺");
                    } else if (val == "\\notin") {
                        node.set_name(name == "mi" ? "mo" : "mml:mo");
                        node.first_child().set_value("∉");
                    } else if (val == "\\ne" || val == "\\neq") {
                        node.set_name(name == "mi" ? "mo" : "mml:mo");
                        node.first_child().set_value("≠");
                    } else if (val == "\\imath") {
                        node.first_child().set_value("ı");
                    } else if (val == "\\jmath") {
                        node.first_child().set_value("ȷ");
                    } else if (val == "\\thetasym") {
                        node.first_child().set_value("ϑ");
                    }
                }

                if (name.find(':') == std::string::npos) {
                    node.set_name((std::string("mml:") + name).c_str());
                }
            }
            return true;
        }
    } walker;
    doc.traverse(walker);

    auto math = doc.child("mml:math");
    if (math) {
        math.append_attribute("xmlns:mml") = "http://www.w3.org/1998/Math/MathML";
    }

    xml_string_writer writer;
    doc.save(writer, "", pugi::format_raw);
    std::string str = writer.result;

    auto remove_all = [&](const std::string &target) {
        size_t pos = 0;
        while ((pos = str.find(target, pos)) != std::string::npos) {
            str.erase(pos, target.length());
        }
    };

    const char *invisible_seqs[] = {"\xE2\x81\xA1",    "\xE2\x81\xA2",    "\xE2\x81\xA3", "\xE2\x81\xA4",
                                    "&#x2061;",        "&#x2062;",        "&#x2063;",     "&#x2064;",
                                    "&ApplyFunction;", "&InvisibleTimes;"};
    for (const char *seq : invisible_seqs) {
        remove_all(seq);
    }

    return str;
}

/**
 * @brief Deep recursive healer for OMML nodes.
 */
void process_omml_node(pugi::xml_node node) {
    std::string name = node.name();
    static const std::set<std::string> containers = {"m:e",   "m:sub", "m:sup", "m:num",
                                                     "m:den", "m:deg", "m:lim", "m:f"};

    // 1. Operator Base Healing: Pull ALL mathematical siblings into empty m:e
    if (name == "m:nary" || name == "m:limLow" || name == "m:limUpp" || name == "m:sSubSup" || name == "m:sSub" ||
        name == "m:sSup") {
        auto e = node.child("m:e");
        if (e && !e.first_child()) {
            pugi::xml_node sib = node.next_sibling();
            while (sib) {
                pugi::xml_node next = sib.next_sibling();
                // Move mathematical elements into the base
                if (sib.type() == pugi::node_element ||
                    (sib.type() == pugi::node_pcdata &&
                     std::string(sib.value()).find_first_not_of(" \t\n\r") != std::string::npos)) {
                    e.append_copy(sib);
                }
                sib.parent().remove_child(sib);
                sib = next;
            }
        }
    }

    // 2. Remove Junk: Strip empty runs (m:r) that contain no text, as they render as boxes
    if (name == "m:r") {
        auto t = node.child("m:t");
        if (!t || std::string(t.child_value()).empty()) {
            pugi::xml_node next = node.next_sibling();
            node.parent().remove_child(node);
            if (next) {
                process_omml_node(next);
            }
            return;
        }
    }

    // 3. hbar mapping fix
    if (name == "m:t") {
        std::string text = node.child_value();
        size_t pos = 0;
        while ((pos = text.find("\xE2\x84\x8F", pos)) != std::string::npos) {
            text.replace(pos, 3, "\xC4\xA7"); // ħ
            node.set_value(text.c_str());
        }
    }

    // 4. Force Fill empty slots with ZWSP
    if (containers.count(name) && !node.first_child()) {
        node.append_child("m:r").append_child("m:t").set_value("\xE2\x80\x8B"); // ZWSP
    }

    for (pugi::xml_node child = node.first_child(); child;) {
        pugi::xml_node next = child.next_sibling();
        process_omml_node(child);
        child = next;
    }
}

std::string sanitize_omml(const std::string &raw_omml) {
    pugi::xml_document doc;
    if (!doc.load_string(raw_omml.c_str())) {
        return raw_omml;
    }
    process_omml_node(doc.root());
    xml_string_writer writer;
    doc.save(writer, "", pugi::format_raw);
    return writer.result;
}

std::string apply_xslt_transformation(const std::string &raw_mathml) {
    std::string mathml = sanitize_mathml(raw_mathml);

    // Parse the XSLT stylesheet directly from the embedded static byte array
    xsltStylesheetPtr cur = nullptr;
    xmlDocPtr style_doc = xmlReadMemory(
        reinterpret_cast<const char*>(openword_resources::MML2OMML_XSL_DATA),
        openword_resources::MML2OMML_XSL_DATA_LEN,
        "MML2OMML.XSL", nullptr, 0);
        
    if (style_doc) {
        cur = xsltParseStylesheetDoc(style_doc);
    }
    if (!cur) {
        return "";
    }
    auto cleanup_xslt = gsl::finally([cur]() { xsltFreeStylesheet(cur); });

    xmlDocPtr xml_doc = xmlReadMemory(mathml.c_str(), static_cast<int>(mathml.length()), nullptr, nullptr, 0);
    if (!xml_doc) {
        return "";
    }
    auto cleanup_doc = gsl::finally([xml_doc]() { xmlFreeDoc(xml_doc); });

    xmlDocPtr res = xsltApplyStylesheet(cur, xml_doc, nullptr);
    if (!res) {
        return "";
    }
    auto cleanup_res = gsl::finally([res]() { xmlFreeDoc(res); });

    xmlChar *xml_res = nullptr;
    int len = 0;
    xsltSaveResultToString(&xml_res, &len, res, cur);
    if (!xml_res) {
        return "";
    }

    std::string omml_str(reinterpret_cast<char *>(xml_res), len);
    xmlFree(xml_res);
    return sanitize_omml(omml_str);
}

} // anonymous namespace

std::string convert_latex_to_omml(const std::string &latex) {
    char *raw_mathml = latex_to_mathml_c(latex.c_str());
    if (!raw_mathml) {
        return "";
    }
    auto cleanup_rust = gsl::finally([raw_mathml]() { free_rust_string(raw_mathml); });
    return apply_xslt_transformation(raw_mathml);
}

std::string convert_mathml_to_omml(const std::string &mathml) { return apply_xslt_transformation(mathml); }

} // namespace openword
