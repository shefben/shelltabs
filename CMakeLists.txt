cmake_minimum_required(VERSION 3.20)
project(ShellTabs LANGUAGES CXX RC)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)
add_compile_definitions(_WIN32_WINNT=0x0A00 NTDDI_VERSION=0x0A000000)
# Global for the whole project
add_definitions(-D_WIN32_WINNT=0x0A00 -DNTDDI_VERSION=0x0A000000)
# Prevent winsock.h conflicts by ensuring winsock2.h is used
add_definitions(-D_WINSOCKAPI_)

option(SHELLTABS_BUILD_TESTS "Build ShellTabs test harnesses" OFF)

add_library(ShellTabs SHARED
    src/dllmain.cpp
    src/ClassFactory.cpp
    src/CExplorerBHO.cpp
    src/ComUtils.cpp
    src/TabBand.cpp
    src/TabBandWindow.cpp
    src/TabManager.cpp
    src/OpenFolderCommand.cpp
    src/BrowserEvents.cpp
    src/FtpClient.cpp
    src/FtpShellFolder.cpp
    src/FtpPidl.cpp
    src/FtpUrl.cpp
    src/Guids.cpp
    src/Module.cpp
    src/SessionStore.cpp
    src/GroupStore.cpp
    src/OptionsStore.cpp
    src/OptionsDialog.cpp
    src/ShellTabsMessages.cpp
    src/Utilities.cpp
    src/Logging.cpp
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
    $<$<CXX_COMPILER_ID:MSVC>:/permissive->
    $<$<CXX_COMPILER_ID:MSVC>:/W4>
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
    dwmapi
    advapi32
    credui
    winhttp
    wininet
    ws2_32
    pathcch
    Shlwapi
    gdiplus
    propsys
)

set_target_properties(ShellTabs PROPERTIES
    OUTPUT_NAME "ShellTabs"
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin
    ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib
    LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin
)

if (WIN32 AND SHELLTABS_BUILD_TESTS)
    add_executable(ShellTabsUtilitiesTests
        tests/UtilitiesFtpTests.cpp
        src/FtpPidl.cpp
        src/FtpUrl.cpp
        src/FtpClient.cpp
        src/Logging.cpp
    )

    target_include_directories(ShellTabsUtilitiesTests PRIVATE
        include
        ${CMAKE_CURRENT_SOURCE_DIR}
    )

    target_compile_definitions(ShellTabsUtilitiesTests PRIVATE
        UNICODE
        _UNICODE
        _WIN32_WINNT=0x0A00
    )

    target_link_libraries(ShellTabsUtilitiesTests PRIVATE
        shlwapi
        urlmon
        shell32
        ole32
        oleaut32
        wininet
        winhttp
        ws2_32
        credui
    )
endif()
