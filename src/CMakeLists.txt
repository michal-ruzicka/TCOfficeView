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

    # Reproducible builds: strip per-build noise from object files and zero the
    # PE TimeDateStamp so the same source always produces the same binary.
    # /Brepro removes the non-deterministic hash MSVC writes into every .obj.
    # /BREPRO (linker) zeroes the PE timestamp field.
    add_compile_options(/Brepro /experimental:deterministic)
    add_link_options(/BREPRO)

    # Normalize absolute source-file paths in debug info so binaries are
    # identical regardless of where the repo is checked out.
    file(TO_NATIVE_PATH "${CMAKE_SOURCE_DIR}/" _src_native)
    add_compile_options("/pathmap:${_src_native}=./")
endif()

# Link the static C/C++ runtime so the artifacts have no vcruntime*.dll
# dependency. Total Commander users expect drop-in install.
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded")

# ---------------------------------------------------------------------------
# Version: version= in pluginst.inf is the single source of truth. It is
# turned into version.h (for the VERSIONINFO resource) at configure time, and
# CMake re-runs whenever pluginst.inf changes.
# ---------------------------------------------------------------------------

set_property(DIRECTORY APPEND PROPERTY CMAKE_CONFIGURE_DEPENDS pluginst.inf)
file(STRINGS pluginst.inf _version_line REGEX "^version=")
# A pre-release suffix (X.Y.Z-devel between releases) is kept in the product
# version string; the numeric FILEVERSION carries MAJOR.MINOR.PATCH.0 only.
if(NOT _version_line MATCHES "^version=(([0-9]+)\\.([0-9]+)\\.([0-9]+)(-[0-9A-Za-z.]+)?)[\r]*$")
    message(FATAL_ERROR "pluginst.inf: no version=MAJOR.MINOR.PATCH[-suffix] line")
endif()
set(TCOV_VERSION "${CMAKE_MATCH_1}")
set(TCOV_VER_MAJOR ${CMAKE_MATCH_2})
set(TCOV_VER_MINOR ${CMAKE_MATCH_3})
set(TCOV_VER_PATCH ${CMAKE_MATCH_4})

# ---------------------------------------------------------------------------
# Plugin DLL: TCOfficeView.wlx (32-bit) or TCOfficeView.wlx64 (64-bit)
# ---------------------------------------------------------------------------

add_library(TCOfficeView SHARED
    TCOfficeView.cpp
    TCOfficeView.def
    TCOfficeView.rc   # VERSIONINFO
)

if(CMAKE_SIZEOF_VOID_P EQUAL 8)
    set(TCOV_SUFFIX ".wlx64")
    set(TCOV_HOST_FILE_NAME "TCOfficeViewHost.exe")
else()
    set(TCOV_SUFFIX ".wlx")
    set(TCOV_HOST_FILE_NAME "TCOfficeViewHost_x86.exe")
endif()
set_target_properties(TCOfficeView PROPERTIES
    SUFFIX "${TCOV_SUFFIX}"
    PREFIX "")

set(TCOV_FILE_NAME "TCOfficeView${TCOV_SUFFIX}")
configure_file(version.h.in version.h @ONLY)
target_include_directories(TCOfficeView PRIVATE ${CMAKE_CURRENT_BINARY_DIR})

# ---------------------------------------------------------------------------
# Host process: TCOfficeViewHost.exe (one per bitness, must match installed Office)
# ---------------------------------------------------------------------------

add_executable(TCOfficeViewHost
    TCOfficeViewHost.cpp
    TCOfficeViewHost.rc   # VERSIONINFO
    app.manifest
)
target_include_directories(TCOfficeViewHost PRIVATE ${CMAKE_CURRENT_BINARY_DIR})

# Build as a Windows GUI subsystem app — no console window flashes when the
# plugin spawns it. wmain is wired up via the entry point property.
set_target_properties(TCOfficeViewHost PROPERTIES
    WIN32_EXECUTABLE TRUE
    LINK_FLAGS "/ENTRY:wmainCRTStartup"
    OUTPUT_NAME "TCOfficeViewHost"
)
