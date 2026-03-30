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
set_target_properties(rust_math_engine PROPERTIES
    IMPORTED_LOCATION "${CMAKE_CURRENT_SOURCE_DIR}/rust_math_engine/target/release/librust_math_engine.a"
)

# Define the OpenWord library
add_library(openword
    src/Document.cpp
    src/Run.cpp
    src/Paragraph.cpp
    src/Table.cpp
    src/Section.cpp
    src/Styles.cpp
    src/Numbering.cpp
    src/MathConverter.cpp
)

target_include_directories(openword
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
