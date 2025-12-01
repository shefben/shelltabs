@echo off
setlocal enableextensions enabledelayedexpansion

rem Determine the directory where the script resides
set "SCRIPT_DIR=%~dp0"
set "DLL_PATH=%SCRIPT_DIR%ShellTabs.dll"

rem Register the DLL silently
regsvr32 /s "%DLL_PATH%"

rem Update the .reg file to point to the current DLL path
set "UPDATED_REG=%TEMP%\ShellTabs_installer.reg"

rem === Pure batch replacement: C:\\ShellTabs.dll -> escaped DLL_PATH ===
set "SEARCH=C:\\ShellTabs.dll"
set "REPLACE=%DLL_PATH:\=\\%"

rem Create a temp file for the modified reg
set "TMP_REG=%UPDATED_REG%.tmp"
if exist "%TMP_REG%" del "%TMP_REG%"

for /f "usebackq delims=" %%A in ("%UPDATED_REG%") do (
    set "LINE=%%A"
    rem Replace occurrences in the current line
    set "LINE=!LINE:%SEARCH%=%REPLACE%!"
    >>"%TMP_REG%" echo(!LINE!
)

rem Overwrite the original file with the updated one
move /y "%TMP_REG%" "%UPDATED_REG%" >nul

rem Import the updated registry file
reg import "%UPDATED_REG%"

endlocal
