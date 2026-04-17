#pragma once

#include "openword/Document.h"
#include <string>

namespace openword {

/**
 * @brief Converts a Word Document object to Markdown text.
 * @param doc The Document to convert.
 * @return A string containing the Markdown representation.
 */
std::string convertToMarkdown(const Document &doc);

} // namespace openword
