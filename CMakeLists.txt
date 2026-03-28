cmake_minimum_required(VERSION 3.15)
project(OpenWord VERSION 0.1.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# Enable testing
enable_testing()

include(FetchContent)

# 1. pugixml (OOXML manipulation) - v1.15 from GitHub
FetchContent_Declare(
    pugixml
    GIT_REPOSITORY https://github.com/zeux/pugixml.git
    GIT_TAG        v1.15
)
FetchContent_MakeAvailable(pugixml)

# 2. zlib-ng (High-performance compression engine for libzip) - v2.3.3
set(ZLIB_COMPAT ON CACHE BOOL "Enable zlib compatibility mode" FORCE)
set(ZLIB_ENABLE_TESTS OFF CACHE BOOL "Disable zlib-ng tests" FORCE)
set(ZLIB_NG_IN_RELOCATABLE_LIB ON CACHE BOOL "Build relocatable library" FORCE)
FetchContent_Declare(
    zlib-ng
    GIT_REPOSITORY https://github.com/zlib-ng/zlib-ng.git
    GIT_TAG        2.3.3
)
FetchContent_MakeAvailable(zlib-ng)

# 3. LibXml2 (MML2OMML transformations) - v2.15.2 from GitHub
set(LIBXML2_WITH_PROGRAMS OFF CACHE BOOL "Disable libxml2 programs" FORCE)
set(LIBXML2_WITH_TESTS OFF CACHE BOOL "Disable libxml2 tests" FORCE)
set(LIBXML2_WITH_PYTHON OFF CACHE BOOL "Disable libxml2 python bindings" FORCE)
set(LIBXML2_WITH_LZMA OFF CACHE BOOL "Disable libxml2 lzma support" FORCE)
set(LIBXML2_WITH_ZLIB OFF CACHE BOOL "Disable libxml2 zlib support" FORCE)
FetchContent_Declare(
    LibXml2
    GIT_REPOSITORY https://github.com/GNOME/libxml2.git
    GIT_TAG        v2.15.2
)
FetchContent_MakeAvailable(LibXml2)

# libxslt (XSLT transformations) - v1.1.45 from GitHub
set(LIBXSLT_WITH_PROGRAMS OFF CACHE BOOL "Disable programs" FORCE)
set(LIBXSLT_WITH_TESTS OFF CACHE BOOL "Disable tests" FORCE)
FetchContent_Declare(
    LibXslt
    GIT_REPOSITORY https://github.com/GNOME/libxslt.git
    GIT_TAG        v1.1.45
)
FetchContent_MakeAvailable(LibXslt)

# 3. libzip (ZIP packaging/unpacking) - v1.11.4 from GitHub
set(BUILD_TOOLS OFF CACHE BOOL "Disable libzip tools" FORCE)
set(BUILD_REGRESS OFF CACHE BOOL "Disable libzip regress tests" FORCE)
set(BUILD_EXAMPLES OFF CACHE BOOL "Disable libzip examples" FORCE)
set(BUILD_DOC OFF CACHE BOOL "Disable libzip documentation" FORCE)
FetchContent_Declare(
    libzip
    GIT_REPOSITORY https://github.com/nih-at/libzip.git
    GIT_TAG        v1.11.4
)
FetchContent_MakeAvailable(libzip)

# 4. fmt (String formatting) - 12.1.0 from GitHub
set(FMT_DOC OFF CACHE BOOL "Disable fmt docs" FORCE)
set(FMT_TEST OFF CACHE BOOL "Disable fmt tests" FORCE)
FetchContent_Declare(
    fmt
    GIT_REPOSITORY https://github.com/fmtlib/fmt.git
    GIT_TAG        12.1.0
)
FetchContent_MakeAvailable(fmt)

# 5. Microsoft GSL (Guidelines Support Library for C++ Safety) - v4.2.1 from GitHub
set(GSL_TEST OFF CACHE BOOL "Disable GSL tests" FORCE)
FetchContent_Declare(
    Microsoft.GSL
    GIT_REPOSITORY https://github.com/microsoft/GSL.git
    GIT_TAG        v4.2.1
)
FetchContent_MakeAvailable(Microsoft.GSL)

# 6. Catch2 (Unit Testing - using v3) - v3.13.0 from GitHub
set(CATCH_INSTALL_DOCS OFF CACHE BOOL "Disable Catch2 docs" FORCE)
set(CATCH_BUILD_TESTING OFF CACHE BOOL "Disable Catch2 testing" FORCE)
FetchContent_Declare(
    Catch2
    GIT_REPOSITORY https://github.com/catchorg/Catch2.git
    GIT_TAG        v3.13.0
)
FetchContent_MakeAvailable(Catch2)
list(APPEND CMAKE_MODULE_PATH ${catch2_SOURCE_DIR}/extras)

# Define the OpenWord library
add_library(openword
    src/Document.cpp
    src/Run.cpp
    src/Paragraph.cpp
    src/Table.cpp
)

# Setup include directories
target_include_directories(openword
    PUBLIC 
        $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
        $<INSTALL_INTERFACE:include>
    PRIVATE
        src
)

# Link dependencies
target_link_libraries(openword
    PUBLIC
        pugixml
        Microsoft.GSL::GSL
    PRIVATE
        LibXml2
        LibXslt
        zip
        fmt::fmt
)

# Add an executable for demonstration
add_executable(openword_demo src/main.cpp)
target_link_libraries(openword_demo PRIVATE openword fmt::fmt)

# Add test using Catch2 v3
add_executable(openword_test tests/test_document.cpp)
target_link_libraries(openword_test PRIVATE openword Catch2::Catch2WithMain zip)

# Enable Catch2 test discovery
include(Catch)
catch_discover_tests(openword_test)
