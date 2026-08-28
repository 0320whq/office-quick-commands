@echo off
cls
echo ========================================
echo   QuickCommands VBA Ribbon 安装脚本
echo ========================================
echo.

set "SCRIPT_DIR=%~dp0"
set "OFFICE_VER=16.0"
set "WORD_STARTUP=%USERPROFILE%\AppData\Roaming\Microsoft\Word\STARTUP"
set "EXCEL_STARTUP=%USERPROFILE%\AppData\Roaming\Microsoft\Excel\XLSTART"
set "TEMPLATES_DIR=%USERPROFILE%\AppData\Roaming\Microsoft\Templates"

if not exist "%WORD_STARTUP%" mkdir "%WORD_STARTUP%"
if not exist "%EXCEL_STARTUP%" mkdir "%EXCEL_STARTUP%"

:: ---- [1/3] Configure macro security ----
echo [1/3] Configuring macro security...
reg add "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Word\Security" /v VBAWarnings /t REG_DWORD /d 2 /f >nul 2>&1
echo   OK Word: macros enabled (warn)
reg add "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Excel\Security" /v VBAWarnings /t REG_DWORD /d 2 /f >nul 2>&1
echo   OK Excel: macros enabled (warn)

reg add "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Word\Security\Trusted Locations\LocQuickCmd" /v Path /t REG_SZ /d "%TEMPLATES_DIR%" /f >nul 2>&1
echo   OK Word trust location

reg add "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Excel\Security\Trusted Locations\LocQuickCmd" /v Path /t REG_SZ /d "%EXCEL_STARTUP%" /f >nul 2>&1
echo   OK Excel trust location

:: ---- [2/3] Copy shell templates to STARTUP ----
echo.
echo [2/3] Copying shell templates to STARTUP...
copy /Y "%SCRIPT_DIR%QuickCommands.dotm" "%WORD_STARTUP%\" >nul 2>&1
if %errorlevel% neq 0 (
    echo   FAIL QuickCommands.dotm, copy manually to:
    echo      %WORD_STARTUP%
) else (
    echo   OK Word: %WORD_STARTUP%\QuickCommands.dotm
)

copy /Y "%SCRIPT_DIR%QuickCommands.xlam" "%EXCEL_STARTUP%\" >nul 2>&1
if %errorlevel% neq 0 (
    echo   FAIL QuickCommands.xlam, copy manually to:
    echo      %EXCEL_STARTUP%
) else (
    echo   OK Excel: %EXCEL_STARTUP%\QuickCommands.xlam
)

:: ---- [3/3] Guide user to import VBA ----
echo.
echo [3/3] Import VBA macros (one-time, ~1 minute):
echo.
echo   === Word ===
echo   1. Open Word (blank document)
echo   2. Press Alt+F11 to open VBA editor
echo   3. Find Normal project on left, right-click - Import File
echo   4. Select: %SCRIPT_DIR%word.bas
echo   5. Press Ctrl+S, close Word
echo   6. Reopen Word, Ribbon tab [QuickCommands] appears
echo.
echo   === Excel ===
echo   1. Open Excel (blank workbook)
echo   2. Press Alt+F11
echo   3. If no VBAProject(PERSONAL.XLSB):
echo      View - Macro - Record macro, save in Personal Macro Workbook
echo   4. Right-click VBAProject(PERSONAL.XLSB) - Import File
echo   5. Select: %SCRIPT_DIR%excel.bas
echo   6. Press Ctrl+S, close Excel
echo   7. Reopen Excel, Ribbon tab [QuickCommands] appears
echo.

echo Open Word now to import macros? [Y=Open Word, N=Later]
choice /c YN /m "Y=Open Word, N=Later"
if %errorlevel%==1 (
    start "" winword.exe
    echo Word opened. Follow steps 2-5 above.
)

echo.
echo Script configuration complete. Press any key to exit.
pause >nul
exit /b 0
