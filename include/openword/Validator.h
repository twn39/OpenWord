#pragma once
#include <gsl/gsl>
#include <memory>
#include <string>

namespace openword {

class SchemaValidator {
  public:
    /**
     * @brief Loads an XSD schema for validation.
     * @param xsdPath Path to the .xsd file.
     */
    explicit SchemaValidator(gsl::czstring xsdPath);
    ~SchemaValidator();

    SchemaValidator(const SchemaValidator &) = delete;
    SchemaValidator &operator=(const SchemaValidator &) = delete;

    /**
     * @brief Checks if the schema was successfully loaded and parsed.
     */
    bool isValid() const;

    /**
     * @brief Validates an XML string against the loaded schema.
     * @param xmlContent The XML string to validate.
     * @param outErrors A string to hold any validation error messages.
     * @return true if valid, false if invalid.
     */
    bool validate(gsl::czstring xmlContent, std::string &outErrors) const;

  private:
    struct Impl;
    std::unique_ptr<Impl> pimpl;
};

} // namespace openword
