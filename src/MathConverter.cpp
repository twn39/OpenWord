#include "Internal.h"
#include <libxml/parser.h>
#include <libxml/tree.h>
#include <libxslt/transform.h>
#include <libxslt/xsltutils.h>
#include <gsl/gsl>
#include <string>

// C-linkage to Rust functions
extern "C" {
    char* latex_to_mathml_c(const char* latex);
    void free_rust_string(char* s);
}

namespace openword {

namespace {

/**
 * @brief Internal helper to apply XSLT transformation from MathML to OMML.
 */
std::string apply_xslt_transformation(const std::string& mathml) {
    // Note: Paths are relative to the execution directory.
    // In production, consider embedding this or resolving the path dynamically.
    gsl::czstring stylesheet_path = "resources/MML2OMML.XSL";
    
    xsltStylesheetPtr cur = xsltParseStylesheetFile(reinterpret_cast<const xmlChar*>(stylesheet_path));
    if (!cur) {
        return "";
    }
    auto cleanup_xslt = gsl::finally([cur]() {
        xsltFreeStylesheet(cur);
    });

    // Parse the MathML XML from memory
    xmlDocPtr doc = xmlReadMemory(mathml.c_str(), static_cast<int>(mathml.length()), nullptr, nullptr, 0);
    if (!doc) {
        return "";
    }
    auto cleanup_doc = gsl::finally([doc]() {
        xmlFreeDoc(doc);
    });

    // Apply the transformation
    xmlDocPtr res = xsltApplyStylesheet(cur, doc, nullptr);
    if (!res) {
        return "";
    }
    auto cleanup_res = gsl::finally([res]() {
        xmlFreeDoc(res);
    });

    // Save the transformation result to a string
    xmlChar* xml_res = nullptr;
    int len = 0;
    xsltSaveResultToString(&xml_res, &len, res, cur);
    
    if (!xml_res) {
        return "";
    }
    
    std::string omml(reinterpret_cast<char*>(xml_res), len);
    xmlFree(xml_res);

    return omml;
}

} // anonymous namespace

std::string convert_latex_to_omml(const std::string& latex) {
    // 1. Rust: LaTeX -> MathML
    char* raw_mathml = latex_to_mathml_c(latex.c_str());
    if (!raw_mathml) {
        return "";
    }
    
    // Ensure the raw Rust string is freed before we return
    auto cleanup_rust = gsl::finally([raw_mathml]() {
        free_rust_string(raw_mathml);
    });

    std::string mathml(raw_mathml);

    // 2. LibXSLT: MathML -> OMML
    return apply_xslt_transformation(mathml);
}

std::string convert_mathml_to_omml(const std::string& mathml) {
    return apply_xslt_transformation(mathml);
}

} // namespace openword
