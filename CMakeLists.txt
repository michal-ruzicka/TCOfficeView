cmake_minimum_required(VERSION 3.20)
project(tcoffice LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Plugin musí být WLX (Total Commander přípona Lister pluginů)
add_library(tcoffice SHARED
    tcoffice.cpp
    tcoffice.def
)

# Statické runtime - plugin nemá závislosti
set_property(TARGET tcoffice PROPERTY
    MSVC_RUNTIME_LIBRARY "MultiThreaded")

# Výstupní jméno: tcoffice.wlx (32-bit) nebo tcoffice.wlx64 (64-bit)
if(CMAKE_SIZEOF_VOID_P EQUAL 8)
    set_target_properties(tcoffice PROPERTIES
        SUFFIX ".wlx64"
        PREFIX "")
else()
    set_target_properties(tcoffice PROPERTIES
        SUFFIX ".wlx"
        PREFIX "")
endif()

target_compile_definitions(tcoffice PRIVATE
    UNICODE _UNICODE
    WIN32_LEAN_AND_MEAN
    NOMINMAX
)
