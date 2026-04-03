cmake_minimum_required(VERSION 3.15)
project(OpenWord VERSION 0.1.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# Enable testing
enable_testing()

include(FetchContent)

# 1. pugixml
FetchContent_Declare(
    pugixml
    GIT_REPOSITORY https://github.com/zeux/pugixml.git
    GIT_TAG        v1.15
)
FetchContent_MakeAvailable(pugixml)

# 2. zlib-ng
set(ZLIB_COMPAT ON CACHE BOOL "" FORCE)
set(ZLIB_ENABLE_TESTS OFF CACHE BOOL "" FORCE)
FetchContent_Declare(
    zlib-ng
    GIT_REPOSITORY https://github.com/zlib-ng/zlib-ng.git
    GIT_TAG        2.3.3
)
FetchContent_MakeAvailable(zlib-ng)

# 3. LibXml2 & LibXslt
find_package(LibXml2 REQUIRED)
find_package(LibXslt REQUIRED)

# 4. libzip
FetchContent_Declare(
    libzip
    GIT_REPOSITORY https://github.com/nih-at/libzip.git
    GIT_TAG        v1.11.4
)
FetchContent_MakeAvailable(libzip)

# 5. fmt
FetchContent_Declare(
    fmt
    GIT_REPOSITORY https://github.com/fmtlib/fmt.git
    GIT_TAG        12.1.0
)
FetchContent_MakeAvailable(fmt)

# 6. Microsoft GSL
FetchContent_Declare(
    Microsoft.GSL
    GIT_REPOSITORY https://github.com/microsoft/GSL.git
    GIT_TAG        v4.2.1
)
FetchContent_MakeAvailable(Microsoft.GSL)

# 7. Catch2
set(CATCH_INSTALL_DOCS OFF CACHE BOOL "" FORCE)
set(CATCH_BUILD_TESTING OFF CACHE BOOL "" FORCE)
FetchContent_Declare(
    Catch2
    GIT_REPOSITORY https://github.com/catchorg/Catch2.git
    GIT_TAG        v3.13.0
)
FetchContent_MakeAvailable(Catch2)
list(APPEND CMAKE_MODULE_PATH ${catch2_SOURCE_DIR}/extras)

# 8. Manual Rust Integration
add_library(rust_math_engine STATIC IMPORTED)

if(MSVC)
    set(RUST_LIB_NAME "rust_math_engine.lib")
else()
    set(RUST_LIB_NAME "librust_math_engine.a")
endif()

set_target_properties(rust_math_engine PROPERTIES
    IMPORTED_LOCATION "${CMAKE_CURRENT_SOURCE_DIR}/rust_math_engine/target/release/${RUST_LIB_NAME}"
)

# Define the OpenWord library

# --- Embed Resources (XSLT) ---
set(MML2OMML_XSL ${CMAKE_CURRENT_SOURCE_DIR}/resources/MML2OMML.XSL)
set(MML2OMML_HEADER ${CMAKE_CURRENT_BINARY_DIR}/generated/MML2OMML_XSL.h)

file(MAKE_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}/generated)

add_custom_command(
    OUTPUT ${MML2OMML_HEADER}
    COMMAND ${CMAKE_COMMAND} 
        -DRESOURCE_IN=${MML2OMML_XSL} 
        -DVARIABLE_NAME=MML2OMML_XSL_DATA 
        -DHEADER_OUT=${MML2OMML_HEADER} 
        -P ${CMAKE_CURRENT_SOURCE_DIR}/cmake/EmbedResource.cmake
    DEPENDS ${MML2OMML_XSL} ${CMAKE_CURRENT_SOURCE_DIR}/cmake/EmbedResource.cmake
    COMMENT "Embedding MML2OMML.XSL into C++ header"
)

add_library(openword
    ${MML2OMML_HEADER}
    src/Document.cpp
    src/Run.cpp
    src/Paragraph.cpp
    src/Table.cpp
    src/Section.cpp
    src/Styles.cpp
    src/Numbering.cpp
    src/MathConverter.cpp
    src/Validator.cpp
    src/HtmlConverter.cpp


)

target_include_directories(openword
    PRIVATE ${CMAKE_CURRENT_BINARY_DIR}/generated
    PUBLIC 
        $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
        $<INSTALL_INTERFACE:include>
    PRIVATE
        src
)

target_link_libraries(openword
    PUBLIC
        pugixml
        Microsoft.GSL::GSL
    PRIVATE
        LibXml2::LibXml2
        LibXslt::LibXslt
        zip
        fmt::fmt
        rust_math_engine
)

if(MSVC)
    # Rust standard library on Windows requires these system libraries
    target_link_libraries(openword PRIVATE Bcrypt ws2_32 user32 advapi32 ntdll userenv)
endif()

# Demo and Tests
add_executable(openword_demo demo/main.cpp)
target_link_libraries(openword_demo PRIVATE openword fmt::fmt)

add_executable(openword_test tests/test_document.cpp)
target_link_libraries(openword_test PRIVATE openword Catch2::Catch2WithMain zip)
include(Catch)
catch_discover_tests(openword_test)
target_sources(openword_test PRIVATE tests/test_extract.cpp)
target_sources(openword_test PRIVATE tests/test_mutability.cpp)
target_sources(openword_test PRIVATE tests/test_metadata.cpp)
target_sources(openword_test PRIVATE tests/test_styles.cpp)
target_sources(openword_test PRIVATE tests/test_header_footer.cpp)
target_sources(openword_test PRIVATE tests/test_validator.cpp)
target_sources(openword_test PRIVATE tests/test_template.cpp)
target_sources(openword_test PRIVATE tests/test_html.cpp)




# Benchmarks
add_executable(openword_benchmark benchmarks/bench_document.cpp)
target_link_libraries(openword_benchmark PRIVATE openword Catch2::Catch2WithMain zip)

add_custom_target(run_benchmark
    COMMAND openword_benchmark
    DEPENDS openword_benchmark
    WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}
    COMMENT "Running OpenWord performance benchmarks..."
)


# --- Doxygen Documentation ---
find_package(Doxygen OPTIONAL_COMPONENTS dot)
if(DOXYGEN_FOUND)
    set(DOXYGEN_IN ${CMAKE_CURRENT_SOURCE_DIR}/Doxyfile.in)
    set(DOXYGEN_OUT ${CMAKE_CURRENT_BINARY_DIR}/Doxyfile)
    configure_file(${DOXYGEN_IN} ${DOXYGEN_OUT} @ONLY)
    
    add_custom_target(doc_doxygen ALL
        COMMAND ${DOXYGEN_EXECUTABLE} ${DOXYGEN_OUT}
        WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}
        COMMENT "Generating API documentation with Doxygen"
        VERBATIM
    )
else()
    message(STATUS "Doxygen need to be installed to generate the doxygen documentation")
endif()
add_executable(openword_template_demo demo/generate_test_doc.cpp)
target_link_libraries(openword_template_demo PRIVATE openword)
add_executable(openword_html_demo demo/generate_html_demo.cpp)
target_link_libraries(openword_html_demo PRIVATE openword)
