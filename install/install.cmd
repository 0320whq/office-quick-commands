@echo off
cls
echo ========================================
echo   快捷指令 QuickCommands 安装脚本
echo ========================================
echo.

:: ---- 检测 Office 版本 ----
set "OFFICE_VER=16.0"
if exist "C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE" (
    set "OFFICE_DIR=C:\Program Files\Microsoft Office\Root\Office16"
) else if exist "C:\Program Files (x86)\Microsoft Office\Root\Office16\WINWORD.EXE" (
    set "OFFICE_DIR=C:\Program Files (x86)\Microsoft Office\Root\Office16"
) else (
    echo [警告] 未检测到 Office 2016/2019/365 安装
    echo 请确认 Office 已正确安装, 然后手动按 README.md 步骤操作
    pause
    exit /b 1
)

set "SCRIPT_DIR=%~dp0"

:: ---- [1/3] 配置宏安全设置 ----
echo [1/3] 配置宏安全设置...

:: Word: 启用宏（级别2 = 警告但不禁用）
reg add "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Word\Security" /v VBAWarnings /t REG_DWORD /d 2 /f >nul 2>&1
echo   [OK] Word 宏安全: 允许运行（带警告）

:: Excel: 同上
reg add "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Excel\Security" /v VBAWarnings /t REG_DWORD /d 2 /f >nul 2>&1
echo   [OK] Excel 宏安全: 允许运行（带警告）

:: 信任 Normal 模板所在目录
set "TEMPLATES_DIR=%USERPROFILE%\AppData\Roaming\Microsoft\Templates"
reg add "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Word\Security\Trusted Locations\LocQuickCmd" /v Path /t REG_SZ /d "%TEMPLATES_DIR%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Word\Security\Trusted Locations\LocQuickCmd" /v AllowSubfolders /t REG_DWORD /d 1 /f >nul 2>&1
echo   [OK] Word 信任位置已添加

set "XLSTART_DIR=%USERPROFILE%\AppData\Roaming\Microsoft\Excel\XLSTART"
if not exist "%XLSTART_DIR%" mkdir "%XLSTART_DIR%"
reg add "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Excel\Security\Trusted Locations\LocQuickCmd" /v Path /t REG_SZ /d "%XLSTART_DIR%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\%OFFICE_VER%\Excel\Security\Trusted Locations\LocQuickCmd" /v AllowSubfolders /t REG_DWORD /d 1 /f >nul 2>&1
echo   [OK] Excel 信任位置已添加

:: ---- [2/3] 准备文件 ----
echo.
echo [2/3] 准备 VBA 源码文件...
echo   Word  源码: %SCRIPT_DIR%word.bas
echo   Excel 源码: %SCRIPT_DIR%excel.bas
echo   [OK] 文件就绪

:: ---- [3/3] 引导用户操作 ----
echo.
echo [3/3] 接下来请手动完成以下步骤（约1分钟）：
echo.
echo   === Word 安装 ===
echo   1. 打开 Word（空白文档即可）
echo   2. 按 Alt+F11 打开 VBA 编辑器
echo   3. 左侧找到 "Normal" 项目, 右键 - 导入文件
echo   4. 选择: %SCRIPT_DIR%word.bas
echo   5. 按 Ctrl+S 保存
echo   6. 关闭 Word
echo.
echo   === Excel 安装 ===
echo   1. 打开 Excel（空白工作簿即可）
echo   2. 按 Alt+F11 打开 VBA 编辑器
echo   3. 如果左侧没有 VBAProject(PERSONAL.XLSB):
echo      - 先在 Excel 里: 视图 - 宏 - 录制宏
echo      - "保存在"选"个人宏工作簿" - 点确定
echo      - 马上点"停止录制"（不用操作任何东西）
echo      - 重新按 Alt+F11, VBAProject(PERSONAL.XLSB) 就出现了
echo   4. 右键 VBAProject(PERSONAL.XLSB) - 导入文件
echo   5. 选择: %SCRIPT_DIR%excel.bas
echo   6. 按 Ctrl+S 保存
echo   7. 关闭 Excel
echo.
echo   === 完成 ===
echo   重新打开 Word 和 Excel, 顶部会出现 "QuickCommands" 工具栏
echo   点击工具栏按钮即可使用各种快捷指令
echo.
echo ========================================
echo  是否现在打开 Word 开始安装?
echo ========================================
choice /c yn /m "Y=打开 Word, N=稍后手动"
if errorlevel 2 goto :done
if errorlevel 1 (
    echo 正在打开 Word...
    start "" winword.exe
)

:done
echo.
echo 安装脚本配置部分已完成。
echo 如需帮助请查看 README.md
echo.
pause
exit /b 0
