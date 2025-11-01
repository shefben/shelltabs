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
    src/PreviewCache.cpp
    src/PreviewOverlay.cpp
    src/IconCache.cpp
    src/OpenFolderCommand.cpp
    src/BrowserEvents.cpp
    src/BackgroundCache.cpp
    src/FtpClient.cpp
    src/FtpShellFolder.cpp
    src/FtpPidl.cpp
    src/FtpUrl.cpp
    src/Guids.cpp
    src/Module.cpp
    src/ColorSerialization.cpp
    src/SessionStore.cpp
    src/GroupStore.cpp
    src/OptionsStore.cpp
    src/OptionsDialog.cpp
    src/ShellTabsMessages.cpp
    src/StringUtils.cpp
    src/Utilities.cpp
    src/Logging.cpp
    src/Notifications.cpp
    src/GlowRenderer.cpp
    src/PaneHooks.cpp
    src/ThemeNotifier.cpp
    src/ExplorerThemeUtils.cpp
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
    $<$<BOOL:${SHELLTABS_BUILD_TESTS}>:SHELLTABS_ENABLE_THEME_TEST_HOOKS>
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
    runtimeobject
    wtsapi32
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

    add_executable(ShellTabsColorSerializationTests
        tests/ColorSerializationTests.cpp
        src/ColorSerialization.cpp
    )

    target_include_directories(ShellTabsColorSerializationTests PRIVATE
        include
        ${CMAKE_CURRENT_SOURCE_DIR}
    )

    target_compile_definitions(ShellTabsColorSerializationTests PRIVATE
        UNICODE
        _UNICODE
        _WIN32_WINNT=0x0A00
    )

    add_executable(ShellTabsPaneHooksTests
        tests/PaneHooksTests.cpp
        src/PaneHooks.cpp
    )

    target_include_directories(ShellTabsPaneHooksTests PRIVATE
        include
        ${CMAKE_CURRENT_SOURCE_DIR}
    )

    add_executable(ShellTabsThemeTransitionTests
        tests/ThemeTransitionTests.cpp
        tests/TestLoggingStubs.cpp
        src/ThemeNotifier.cpp
    )

    target_include_directories(ShellTabsThemeTransitionTests PRIVATE
        include
        ${CMAKE_CURRENT_SOURCE_DIR}
    )

    add_executable(ShellTabsTabManagerTests
        tests/TabManagerWindowTests.cpp
        src/TabManager.cpp
        src/Utilities.cpp
        src/Logging.cpp
        src/ShellTabsMessages.cpp
        src/StringUtils.cpp
    )

    target_include_directories(ShellTabsTabManagerTests PRIVATE
        include
        ${CMAKE_CURRENT_SOURCE_DIR}
    )

    target_compile_definitions(ShellTabsPaneHooksTests PRIVATE
        UNICODE
        _UNICODE
        _WIN32_WINNT=0x0A00
    )

    target_link_libraries(ShellTabsPaneHooksTests PRIVATE
        user32
        comctl32
    )

    target_compile_definitions(ShellTabsThemeTransitionTests PRIVATE
        UNICODE
        _UNICODE
        _WIN32_WINNT=0x0A00
        SHELLTABS_ENABLE_THEME_TEST_HOOKS
    )

    target_link_libraries(ShellTabsThemeTransitionTests PRIVATE
        runtimeobject
        wtsapi32
        shlwapi
        shell32
        ole32
        oleaut32
        uuid
    )

    add_executable(ShellTabsTabBandWindowDiffTests
        tests/TabBandWindowDiffTests.cpp
    )

    target_include_directories(ShellTabsTabBandWindowDiffTests PRIVATE
        include
        ${CMAKE_CURRENT_SOURCE_DIR}
    )

    target_compile_definitions(ShellTabsTabBandWindowDiffTests PRIVATE
        UNICODE
        _UNICODE
        _WIN32_WINNT=0x0A00
        SHELLTABS_BUILD_TESTS
    )

    target_link_libraries(ShellTabsTabBandWindowDiffTests PRIVATE
        ShellTabs
        shlwapi
        shell32
        ole32
        oleaut32
        comctl32
    )

    add_executable(ShellTabsGuardTests
        tests/UtilitiesGuardTests.cpp
    )

    target_include_directories(ShellTabsGuardTests PRIVATE
        include
        ${CMAKE_CURRENT_SOURCE_DIR}
    )

    target_compile_definitions(ShellTabsGuardTests PRIVATE
        UNICODE
        _UNICODE
        _WIN32_WINNT=0x0A00
    )
endif()
