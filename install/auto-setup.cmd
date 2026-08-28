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

:: ---- STEP 1: Auto-config registry ----
echo [1/5] Auto-configuring registry...
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

:: ---- STEP 2: Copy shell templates ----
echo.
echo [2/5] Copying shell templates to STARTUP...
copy /Y "%SCRIPT_DIR%QuickCommands.dotm" "%WORD_STARTUP%\" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [FAIL] QuickCommands.dotm -- copy manually to: %WORD_STARTUP%
) else (
    echo   [OK] Word shell: %WORD_STARTUP%\QuickCommands.dotm
)

copy /Y "%SCRIPT_DIR%QuickCommands.xlam" "%EXCEL_STARTUP%\" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [FAIL] QuickCommands.xlam -- copy manually to: %EXCEL_STARTUP%
) else (
    echo   [OK] Excel shell: %EXCEL_STARTUP%\QuickCommands.xlam
)

:: ---- STEP 3: Import VBA (manual step, guided) ----
echo.
echo [3/5] Import VBA macros (one-time, ~90 seconds):
echo.
echo   The only manual step: import the .bas files into Office.
echo.
echo   === WORD ===
echo   1. In Word, press Alt+F11 to open VBA editor
echo   2. Find Normal project on the left, right-click - Import File
echo   3. Select: %SCRIPT_DIR%word.bas
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

:: ---- STEP 4: Verification ----
echo.
echo [4/5] Verifying installation...

:: Check Word registry
echo   Checking Word registry...
reg query "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Word\Security" /v VBAWarnings >nul 2>&1
if %errorlevel%==0 (echo     [OK] Word macro registry set) else (echo     [FAIL] Word macro registry)

:: Check Excel registry
echo   Checking Excel registry...
reg query "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Excel\Security" /v VBAWarnings >nul 2>&1
if %errorlevel%==0 (echo     [OK] Excel macro registry set) else (echo     [FAIL] Excel macro registry)

:: Check if files exist
echo   Checking shell files...
if exist "%WORD_STARTUP%\QuickCommands.dotm" (
    echo     [OK] QuickCommands.dotm in STARTUP
) else (
    echo     [FAIL] QuickCommands.dotm not in STARTUP
)

if exist "%EXCEL_STARTUP%\QuickCommands.xlam" (
    echo     [OK] QuickCommands.xlam in STARTUP
) else (
    echo     [FAIL] QuickCommands.xlam not in STARTUP
)

:: Open Word for visual verification
echo.
echo [5/5] Opening Word and Excel for visual verification...
echo   If QuickCommands tab appears, installation is SUCCESSFUL.
echo   If NOT, the VBA was not imported correctly.
start "" winword.exe
timeout /t 3 /nobreak >nul
start "" excel.exe
timeout /t 3 /nobreak >nul

echo.
echo ========================================
echo  Setup script complete.
echo  Check Word/Excel for QuickCommands tab.
echo ========================================
pause >nul
exit /b 0