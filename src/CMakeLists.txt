cmake_minimum_required(VERSION 3.20)
project(TCOfficeView LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Shared settings for every target in this project.
add_compile_definitions(
    UNICODE _UNICODE
    WIN32_LEAN_AND_MEAN
    NOMINMAX
)

# Source files are UTF-8 (no BOM). Without /utf-8 MSVC interprets non-ASCII
# bytes in the source using the system ANSI code page, which produces
# mojibake in L"..." wide string literals on machines whose ACP isn't UTF-8.
# /utf-8 is equivalent to /source-charset:utf-8 /execution-charset:utf-8.
if(MSVC)
    add_compile_options(/utf-8)
endif()

# Link the static C/C++ runtime so the artifacts have no vcruntime*.dll
# dependency. Total Commander users expect drop-in install.
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded")

# ---------------------------------------------------------------------------
# Plugin DLL: TCOfficeView.wlx (32-bit) or TCOfficeView.wlx64 (64-bit)
# ---------------------------------------------------------------------------

add_library(TCOfficeView SHARED
    TCOfficeView.cpp
    TCOfficeView.def
)

if(CMAKE_SIZEOF_VOID_P EQUAL 8)
    set_target_properties(TCOfficeView PROPERTIES
        SUFFIX ".wlx64"
        PREFIX "")
else()
    set_target_properties(TCOfficeView PROPERTIES
        SUFFIX ".wlx"
        PREFIX "")
endif()

# ---------------------------------------------------------------------------
# Host process: TCOfficeViewHost.exe (one per bitness, must match installed Office)
# ---------------------------------------------------------------------------

add_executable(TCOfficeViewHost
    TCOfficeViewHost.cpp
    app.manifest
)

# Build as a Windows GUI subsystem app — no console window flashes when the
# plugin spawns it. wmain is wired up via the entry point property.
set_target_properties(TCOfficeViewHost PROPERTIES
    WIN32_EXECUTABLE TRUE
    LINK_FLAGS "/ENTRY:wmainCRTStartup"
    OUTPUT_NAME "TCOfficeViewHost"
)
