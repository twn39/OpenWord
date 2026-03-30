# Basic script to convert a file into a C/C++ static byte array
function(embed_resource RESOURCE_FILE VARIABLE_NAME OUTPUT_HEADER)
    file(READ ${RESOURCE_FILE} HEX_CONTENT HEX)
    string(REGEX MATCHALL "([a-f0-9][a-f0-9])" HEX_DATA ${HEX_CONTENT})
    
    set(ARRAY_CONTENT "")
    set(COUNTER 0)
    foreach(HEX_BYTE ${HEX_DATA})
        string(APPEND ARRAY_CONTENT "0x${HEX_BYTE}, ")
        math(EXPR COUNTER "${COUNTER}+1")
        if (COUNTER EQUAL 16)
            string(APPEND ARRAY_CONTENT "\n    ")
            set(COUNTER 0)
        endif()
    endforeach()
    
    file(WRITE ${OUTPUT_HEADER} "#pragma once\n\n")
    file(APPEND ${OUTPUT_HEADER} "namespace openword_resources {\n\n")
    file(APPEND ${OUTPUT_HEADER} "constexpr unsigned char ${VARIABLE_NAME}[] = {\n    ${ARRAY_CONTENT}0x00\n};\n\n")
    file(APPEND ${OUTPUT_HEADER} "constexpr unsigned int ${VARIABLE_NAME}_LEN = sizeof(${VARIABLE_NAME}) - 1;\n\n")
    file(APPEND ${OUTPUT_HEADER} "}\n")
endfunction()

embed_resource("${RESOURCE_IN}" "${VARIABLE_NAME}" "${HEADER_OUT}")
