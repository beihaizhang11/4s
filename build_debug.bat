@echo off
chcp 65001 >nul
echo ========================================
echo 汽车维修服务账单生成器 - 调试模式打包
echo ========================================
echo.
echo 注意: 此模式会显示控制台窗口，用于调试
echo.

echo [1/3] 清理旧文件...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "*.spec" del /q "*.spec"
echo ✓ 清理完成

echo.
echo [2/3] 开始打包 (调试模式)...
pyinstaller --name="汽车维修服务账单生成器_Debug" ^
    --onefile ^
    --console ^
    --hidden-import=win32timezone ^
    --hidden-import=openpyxl.cell._writer ^
    --collect-all=openpyxl ^
    --noconfirm ^
    main.py

if %errorlevel% neq 0 (
    echo ✗ 打包失败
    pause
    exit /b 1
)

echo ✓ 打包完成

echo.
echo [3/3] 完成
echo 可执行文件: dist\汽车维修服务账单生成器_Debug.exe
echo.
explorer "dist"
pause
