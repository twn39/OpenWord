#include <openword/Document.h>
#include <iostream>

int main() {
    openword::Document doc;
    doc.addHtml("<p><b>Bold</b> and <i>Italic</i> with <span style=\"color: #FF0000\">Red</span> text.</p>");
    auto elements = doc.elements();
    std::cout << "Elements size: " << elements.size() << std::endl;
    return 0;
}
