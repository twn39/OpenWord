#include <fmt/core.h>
#include <libxml/parser.h>
#include <libxml/xmlschemas.h>
#include <libxml/xmlschemastypes.h>
#include <openword/Validator.h>

namespace openword {

struct SchemaValidator::Impl {
    xmlSchemaPtr schema = nullptr;

    ~Impl() {
        if (schema) {
            xmlSchemaFree(schema);
        }
    }
};

SchemaValidator::SchemaValidator(gsl::czstring xsdPath) : pimpl(std::make_unique<Impl>()) {
    xmlLineNumbersDefault(1);

    xmlSchemaParserCtxtPtr ctxt = xmlSchemaNewParserCtxt(xsdPath);
    if (ctxt) {
        xmlSchemaSetParserErrors(ctxt, (xmlSchemaValidityErrorFunc)fprintf, (xmlSchemaValidityWarningFunc)fprintf,
                                 stderr);
        pimpl->schema = xmlSchemaParse(ctxt);
        xmlSchemaFreeParserCtxt(ctxt);
    }
}

SchemaValidator::~SchemaValidator() = default;

bool SchemaValidator::isValid() const { return pimpl->schema != nullptr; }

static void validityErrorFunc(void *ctx, const char *msg, ...) {
    std::string *outErrors = static_cast<std::string *>(ctx);
    char buf[1024];
    va_list args;
    va_start(args, msg);
    vsnprintf(buf, sizeof(buf), msg, args);
    va_end(args);
    outErrors->append(buf);
}

bool SchemaValidator::validate(gsl::czstring xmlContent, std::string &outErrors) const {
    outErrors.clear();
    if (!pimpl->schema) {
        outErrors = "Schema not loaded.";
        return false;
    }

    xmlDocPtr doc =
        xmlReadMemory(xmlContent, static_cast<int>(std::string_view(xmlContent).length()), "noname.xml", nullptr, 0);
    if (!doc) {
        outErrors = "Failed to parse XML string.";
        return false;
    }

    xmlSchemaValidCtxtPtr valid_ctxt = xmlSchemaNewValidCtxt(pimpl->schema);
    if (!valid_ctxt) {
        xmlFreeDoc(doc);
        outErrors = "Failed to create validation context.";
        return false;
    }

    xmlSchemaSetValidErrors(valid_ctxt, (xmlSchemaValidityErrorFunc)validityErrorFunc,
                            (xmlSchemaValidityWarningFunc)validityErrorFunc, &outErrors);

    int ret = xmlSchemaValidateDoc(valid_ctxt, doc);

    xmlSchemaFreeValidCtxt(valid_ctxt);
    xmlFreeDoc(doc);

    return (ret == 0);
}

} // namespace openword
