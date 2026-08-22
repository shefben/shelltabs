cmake_minimum_required(VERSION 3.20)
project(ShellTabs LANGUAGES C CXX RC)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)
add_compile_definitions(_WIN32_WINNT=0x0A00 NTDDI_VERSION=0x0A000000)
# Global for the whole project
add_definitions(-D_WIN32_WINNT=0x0A00 -DNTDDI_VERSION=0x0A000000 -DNOMINMAX)
# Prevent winsock.h conflicts by ensuring winsock2.h is used
add_definitions(-D_WINSOCKAPI_)

option(SHELLTABS_BUILD_TESTS "Build ShellTabs test harnesses" OFF)

add_library(minhook STATIC
    third_party/minhook/src/buffer.cpp
    third_party/minhook/src/export.cpp
    third_party/minhook/src/hook.cpp
    third_party/minhook/src/thread.cpp
    third_party/minhook/src/trampoline.cpp
    third_party/minhook/src/HDE64/src/hde64.c
)

target_include_directories(minhook PUBLIC
    third_party/minhook
    third_party/minhook/src
)

target_compile_definitions(minhook PRIVATE
    _CRT_SECURE_NO_WARNINGS
)

set_source_files_properties(src/resource.rc PROPERTIES)

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
    src/ColorUtils.cpp
    src/OpenFolderCommand.cpp
    src/BrowserEvents.cpp
    src/BackgroundCache.cpp
    src/FtpClient.cpp
    src/NamespaceContextMenu.cpp
    src/FtpShellFolder.cpp
    src/FtpPidl.cpp
    src/FtpUrl.cpp
    src/HttpClient.cpp
    src/HttpShellFolder.cpp
    src/HttpPidl.cpp
    src/HttpUrl.cpp
    src/HtmlDirectoryParser.cpp
    src/Guids.cpp
    src/Module.cpp
    src/ColorSerialization.cpp
    src/SessionStore.cpp
    src/CrashRecoveryLogger.cpp
    src/DirectoryWatcher.cpp
    src/GroupStore.cpp
    src/OptionsStore.cpp
    src/OptionsDialog.cpp
    src/ShellTabsMessages.cpp
    src/StringUtils.cpp
    src/Utilities.cpp
    src/Logging.cpp
    src/Notifications.cpp
    src/BreadcrumbGradient.cpp
    src/DpiUtils.cpp
    src/EditGradientRenderer.cpp
    src/ExplorerGlowSurfaces.cpp
    src/CompositionIntercept.cpp
    src/FolderBackgroundHooks.cpp
    src/ThemeHooks.cpp
    src/ThemeNotifier.cpp
    src/ExplorerThemeUtils.cpp
    src/ComVTableHook.cpp
    src/TaskbarPreviewHook.cpp
    src/TaskbarTabListPopup.cpp
    src/ShellTabsTabProvider.cpp
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
    # Generate PDB even for Release builds (for crash diagnostics)
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:Release>>:/Zi>
)

# Generate MAP file and keep PDB for Release builds
target_link_options(ShellTabs PRIVATE
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:Release>>:/MAP>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:Release>>:/DEBUG>
)

target_link_libraries(ShellTabs PRIVATE
    minhook
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
    d2d1
    dwrite
    propsys
    uiautomationcore
    runtimeobject
    wtsapi32
    dbghelp
    gdiplus
    oleacc
)

set_target_properties(ShellTabs PROPERTIES
    OUTPUT_NAME "ShellTabs"
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin
    ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib
    LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin
)

add_executable(ShellTabsWorker src/ShellTabsWorker.cpp)
target_compile_definitions(ShellTabsWorker PRIVATE UNICODE _UNICODE _WIN32_WINNT=0x0A00)
target_link_libraries(ShellTabsWorker PRIVATE ole32 oleaut32 shell32 shlwapi)
set_target_properties(ShellTabsWorker PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin
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
        dbghelp
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

    add_executable(ShellTabsIconCacheTests
        tests/IconCacheStressTests.cpp
        tests/IconCacheTestStubs.cpp
        tests/TestLoggingStubs.cpp
        src/IconCache.cpp
    )

    target_include_directories(ShellTabsIconCacheTests PRIVATE
        include
        ${CMAKE_CURRENT_SOURCE_DIR}
    )

    target_compile_definitions(ShellTabsIconCacheTests PRIVATE
        UNICODE
        _UNICODE
        _WIN32_WINNT=0x0A00
        NOMINMAX
        SHELLTABS_BUILD_TESTS
    )

    target_link_libraries(ShellTabsIconCacheTests PRIVATE
        user32
        shell32
    )

    add_executable(ShellTabsHtmlParserTests
        tests/HtmlDirectoryParserTests.cpp
        src/HtmlDirectoryParser.cpp
    )

    target_include_directories(ShellTabsHtmlParserTests PRIVATE
        include
        ${CMAKE_CURRENT_SOURCE_DIR}
    )

    target_compile_definitions(ShellTabsHtmlParserTests PRIVATE
        UNICODE
        _UNICODE
        _WIN32_WINNT=0x0A00
    )

    target_link_libraries(ShellTabsHtmlParserTests PRIVATE
        shlwapi
        shell32
        ole32
        oleaut32
        dbghelp
    )

    add_executable(ShellTabsHttpPidlTests
        tests/HttpPidlTests.cpp
        tests/HttpPidlTestStubs.cpp
        src/HttpPidl.cpp
        src/HttpUrl.cpp
    )

    target_include_directories(ShellTabsHttpPidlTests PRIVATE
        include
        ${CMAKE_CURRENT_SOURCE_DIR}
    )

    target_compile_definitions(ShellTabsHttpPidlTests PRIVATE
        UNICODE
        _UNICODE
        _WIN32_WINNT=0x0A00
        NOMINMAX
    )

    target_link_libraries(ShellTabsHttpPidlTests PRIVATE
        shlwapi
        urlmon
        shell32
        ole32
        oleaut32
        winhttp
        dbghelp
    )
endif()
