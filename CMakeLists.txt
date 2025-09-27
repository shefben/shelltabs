cmake_minimum_required(VERSION 3.20)
project(ShellTabs LANGUAGES CXX RC)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

add_library(ShellTabs SHARED
    src/dllmain.cpp
    src/ClassFactory.cpp
    src/CExplorerBHO.cpp
    src/ComUtils.cpp
    src/TabBand.cpp
    src/TabBandWindow.cpp
    src/TabManager.cpp
    src/FolderViewColorizer.cpp
    src/BrowserEvents.cpp
    src/Guids.cpp
    src/Module.cpp
    src/TagColumnProvider.cpp
    src/Tagging.cpp
    src/SessionStore.cpp
    src/GroupStore.cpp
    src/Utilities.cpp
    src/GitStatus.cpp
    src/resource.rc
)

if (MSVC)
    target_sources(ShellTabs PRIVATE ShellTabs.def)
endif()

target_include_directories(ShellTabs PRIVATE
    include
    ${CMAKE_CURRENT_SOURCE_DIR}
)

target_compile_definitions(ShellTabs PRIVATE
    UNICODE
    _UNICODE
    _WIN32_WINNT=0x0A00
    NOMINMAX
)

target_compile_options(ShellTabs PRIVATE
    $<$<CXX_COMPILER_ID:MSVC>:/permissive-/W4>
    $<$<CXX_COMPILER_ID:MSVC>:/EHsc>
)

target_link_libraries(ShellTabs PRIVATE
    comctl32
    shlwapi
    shell32
    ole32
    oleaut32
    urlmon
    uuid
    msimg32
    comdlg32
    uxtheme
    advapi32
    pathcch
)

set_target_properties(ShellTabs PROPERTIES
    OUTPUT_NAME "ShellTabs"
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin
    ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib
    LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin
)
