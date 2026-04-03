#include <libxml/HTMLparser.h>
#include <iostream>

void parseHtml(const char* html) {
    htmlDocPtr doc = htmlReadMemory(html, strlen(html), nullptr, "UTF-8", HTML_PARSE_NOERROR | HTML_PARSE_NOWARNING | HTML_PARSE_RECOVER);
    if (!doc) {
        std::cerr << "Failed to parse HTML" << std::endl;
        return;
    }
    
    xmlNodePtr root = xmlDocGetRootElement(doc);
    std::cout << "Root: " << root->name << std::endl;
    
    xmlFreeDoc(doc);
}

int main() {
    parseHtml("<h1>Hello <br> <b>World!</b></h1>");
    return 0;
}
