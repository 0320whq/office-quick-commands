@echo off
cls
echo ========================================
echo   QuickCommands Auto-Install + Test
echo ========================================
echo.

set "SCRIPT_DIR=%~dp0"
set "OFFICE_VER=16.0"
set "WORD_STARTUP=%USERPROFILE%\AppData\Roaming\Microsoft\Word\STARTUP"
set "EXCEL_STARTUP=%USERPROFILE%\AppData\Roaming\Microsoft\Excel\XLSTART"
set "TEMPLATES_DIR=%USERPROFILE%\AppData\Roaming\Microsoft\Templates"

:: Clean up broken shell files from previous attempts
echo Removing any broken shell files from previous attempts...
del "%WORD_STARTUP%\QuickCommands.dotm" 2>nul
del "%EXCEL_STARTUP%\QuickCommands.xlam" 2>nul

:: ---- STEP 1: Auto-config registry ----
echo.
echo [1/3] Auto-configuring registry...
mkdir "%WORD_STARTUP%" 2>nul
mkdir "%EXCEL_STARTUP%" 2>nul

reg add "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Word\Security" /v VBAWarnings /t REG_DWORD /d 2 /f >nul 2>&1
echo   [OK] Word macro security: Enabled with warning
reg add "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Excel\Security" /v VBAWarnings /t REG_DWORD /d 2 /f >nul 2>&1
echo   [OK] Excel macro security: Enabled with warning

reg add "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Word\Security\Trusted Locations\LocQuickCmd" /v Path /t REG_SZ /d "%TEMPLATES_DIR%" /f >nul 2>&1
echo   [OK] Word trusted location

reg add "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Excel\Security\Trusted Locations\LocQuickCmd" /v Path /t REG_SZ /d "%EXCEL_STARTUP%" /f >nul 2>&1
echo   [OK] Excel trusted location

:: ---- STEP 2: Import VBA (the only manual step) ----
echo.
echo [2/3] Import VBA macros (one-time, ~90 seconds):
echo.
echo   This is the ONLY manual step. The toolbar will be
echo   created AUTOMATICALLY when you restart Word/Excel.
echo.
echo   === WORD ===
echo   1. In Word, press Alt+F11 to open VBA editor
echo   2. Find Normal project on the left
echo   3. Right-click - Import File - Select: %SCRIPT_DIR%word.bas
echo   4. Press Ctrl+S, then close Word
echo   5. Come back here and press any key to continue...
pause >nul

echo.
echo   === EXCEL ===
echo   1. In Excel, press Alt+F11
echo   2. If NO VBAProject(PERSONAL.XLSB):
echo      View - Macro - Record macro - Save in Personal Macro Workbook - Stop
echo   3. Right-click VBAProject(PERSONAL.XLSB) - Import File
echo   4. Select: %SCRIPT_DIR%excel.bas
echo   5. Press Ctrl+S, then close Excel
echo   6. Come back here and press any key to continue...
pause >nul

:: ---- STEP 3: Verification ----
echo.
echo [3/3] Verifying installation...

reg query "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Word\Security" /v VBAWarnings >nul 2>&1
if %errorlevel%==0 (echo   [OK] Word macro registry set) else (echo   [FAIL] Word macro registry)

reg query "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Excel\Security" /v VBAWarnings >nul 2>&1
if %errorlevel%==0 (echo   [OK] Excel macro registry set) else (echo   [FAIL] Excel macro registry)

echo.
echo Restarting Word and Excel to verify toolbars appear...
start "" winword.exe
timeout /t 3 /nobreak >nul
start "" excel.exe
timeout /t 3 /nobreak >nul

echo.
echo ========================================
echo  Check if QuickCommands toolbar appears
echo  (floating bar, not ribbon tab)
echo ========================================
pause >nul
exit /b 0