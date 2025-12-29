@echo off
chcp 65001 >nul
echo ========================================
echo 汽车维修服务账单生成器 - 打包工具
echo ========================================
echo.

echo [1/4] 检查依赖...
python -c "import PyQt6; import openpyxl; import win32com.client; import PyInstaller" 2>nul
if %errorlevel% neq 0 (
    echo 错误: 缺少必要的依赖包
    echo 正在安装依赖...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo 安装依赖失败，请手动运行: pip install -r requirements.txt
        pause
        exit /b 1
    )
)
echo ✓ 依赖检查完成

echo.
echo [2/4] 清理旧的打包文件...
if exist "build" (
    rmdir /s /q "build"
    echo ✓ 删除 build 目录
)
if exist "dist" (
    rmdir /s /q "dist"
    echo ✓ 删除 dist 目录
)
if exist "*.spec" (
    del /q "*.spec"
    echo ✓ 删除旧的 spec 文件
)
echo ✓ 清理完成

echo.
echo [3/4] 开始打包...
echo 这可能需要几分钟时间，请耐心等待...
pyinstaller --name="汽车维修服务账单生成器" ^
    --windowed ^
    --onefile ^
    --icon=NONE ^
    --add-data="config.json;." ^
    --hidden-import=win32timezone ^
    --hidden-import=openpyxl.cell._writer ^
    --hidden-import=PyQt6.QtCore ^
    --hidden-import=PyQt6.QtGui ^
    --hidden-import=PyQt6.QtWidgets ^
    --collect-all=openpyxl ^
    --noconfirm ^
    main.py

if %errorlevel% neq 0 (
    echo.
    echo ✗ 打包失败！
    pause
    exit /b 1
)

echo ✓ 打包完成

echo.
echo [4/4] 整理输出文件...
if exist "dist\汽车维修服务账单生成器.exe" (
    echo ✓ 可执行文件已生成: dist\汽车维修服务账单生成器.exe
    
    REM 复制配置文件模板（如果不存在）
    if not exist "dist\config.json" (
        if exist "config.json" (
            copy "config.json" "dist\config.json" >nul
            echo ✓ 已复制配置文件模板
        )
    )
    
    REM 创建output文件夹
    if not exist "dist\output" (
        mkdir "dist\output"
        echo ✓ 已创建output文件夹
    )
    
    echo.
    echo ========================================
    echo 打包成功！
    echo ========================================
    echo.
    echo 可执行文件位置: dist\汽车维修服务账单生成器.exe
    echo.
    echo 使用说明:
    echo 1. 将 dist 文件夹中的内容复制到目标位置
    echo 2. 双击 汽车维修服务账单生成器.exe 运行程序
    echo 3. 首次运行需要配置数据库路径等设置
    echo.
    
    REM 打开dist文件夹
    explorer "dist"
) else (
    echo ✗ 未找到生成的可执行文件
    pause
    exit /b 1
)

echo.
pause
