cmake_minimum_required(VERSION 3.20)
project(tcoffice LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Shared settings for every target in this project.
add_compile_definitions(
    UNICODE _UNICODE
    WIN32_LEAN_AND_MEAN
    NOMINMAX
)

# Link the static C/C++ runtime so the artifacts have no vcruntime*.dll
# dependency. Total Commander users expect drop-in install.
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded")

# ---------------------------------------------------------------------------
# Plugin DLL: tcoffice.wlx (32-bit) or tcoffice.wlx64 (64-bit)
# ---------------------------------------------------------------------------

add_library(tcoffice SHARED
    tcoffice.cpp
    tcoffice.def
)

if(CMAKE_SIZEOF_VOID_P EQUAL 8)
    set_target_properties(tcoffice PROPERTIES
        SUFFIX ".wlx64"
        PREFIX "")
else()
    set_target_properties(tcoffice PROPERTIES
        SUFFIX ".wlx"
        PREFIX "")
endif()

# ---------------------------------------------------------------------------
# Host process: tcoffice_host.exe (one per bitness, must match installed Office)
# ---------------------------------------------------------------------------

add_executable(tcoffice_host
    tcoffice_host.cpp
    app.manifest
)

# Build as a Windows GUI subsystem app — no console window flashes when the
# plugin spawns it. wmain is wired up via the entry point property.
set_target_properties(tcoffice_host PROPERTIES
    WIN32_EXECUTABLE TRUE
    LINK_FLAGS "/ENTRY:wmainCRTStartup"
    OUTPUT_NAME "tcoffice_host"
)
