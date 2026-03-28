#pragma once

#include <string>
#include <memory>
#include <gsl/gsl>

namespace openword {

/**
 * @brief Represents a Run of text within a paragraph (w:r).
 * This is a lightweight proxy object passed by value.
 */
class Run {
public:
    explicit Run(void* node); // Internal use

    Run& setText(const std::string& text);
    Run& setBold(bool val = true);
    Run& setItalic(bool val = true);
    Run& setFontSize(int halfPoints);

private:
    void* node_; 
};

/**
 * @brief Represents a Paragraph (w:p).
 * This is a lightweight proxy object passed by value.
 */
class Paragraph {
public:
    explicit Paragraph(void* node); // Internal use
    
    Run addRun(const std::string& text = "");
    
    Paragraph& setStyle(gsl::czstring styleId);
    Paragraph& setAlignment(gsl::czstring align); // "left", "center", "right", "both"

private:
    void* node_;
};

/**
 * @brief Represents a Microsoft Word Document (.docx)
 */
class Document {
public:
    Document();
    ~Document();

    // Disable copy for safety (RAII)
    Document(const Document&) = delete;
    Document& operator=(const Document&) = delete;

    // --- IO Operations ---
    bool save(gsl::czstring filepath);

    // --- Content Creation ---
    Paragraph addParagraph(const std::string& text = "");
    
    /**
     * @brief Define a basic paragraph style
     * @param styleId Unique ID for the style (e.g., "Heading1")
     * @param name Human-readable name
     */
    void addStyle(gsl::czstring styleId, gsl::czstring name);

    // --- Utilities ---
    std::string convertMathMLToOMML(const std::string& mathml) const;

private:
    struct Impl;
    std::unique_ptr<Impl> pimpl;
};

} // namespace openword
