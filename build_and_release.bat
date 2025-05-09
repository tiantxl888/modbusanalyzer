@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

REM 切换到脚本所在目录
cd /d "%~dp0"

REM 检查Python环境
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python not found in PATH
    pause
    exit /b 1
)

echo 正在更新 pip...
python -m pip install --upgrade pip --user

echo 正在安装必要的包...
python -m pip install --upgrade pyinstaller --user

echo.
echo 开始打包主程序...

REM 清理旧文件
if exist "dist" (
    echo 清理旧的打包文件...
    rd /s /q "dist"
)
if exist "build" (
    echo 清理旧的构建文件...
    rd /s /q "build"
)
if exist "*.spec" (
    echo 清理旧的spec文件...
    del /f /q *.spec
)
if exist "release" (
    echo 清理旧的release目录...
    rd /s /q "release"
)

REM 运行PyInstaller，优化打包选项
python -m PyInstaller --noconfirm --windowed --clean ^
    --exclude-module matplotlib ^
    --exclude-module scipy ^
    --exclude-module sklearn ^
    --exclude-module torch ^
    --exclude-module tensorflow ^
    --exclude-module PIL ^
    --exclude-module cv2 ^
    --add-data "ui;ui" ^
    --add-data "core;core" ^
    --add-data "utils;utils" ^
    --add-data "plugins;plugins" ^
    --add-data "config_and_params.xlsx;." ^
    --hidden-import PyQt5.QtCore ^
    --hidden-import PyQt5.QtGui ^
    --hidden-import PyQt5.QtWidgets ^
    --hidden-import pyqtgraph ^
    --hidden-import pandas ^
    --hidden-import numpy ^
    --hidden-import pyserial ^
    --hidden-import openpyxl ^
    modbus_analyzer.py

if %errorlevel% neq 0 (
    echo Error: 打包失败
    echo 请检查错误信息
    pause
    exit /b 1
)

REM 获取版本号
set "version=1.0.0"
for /f "tokens=*" %%a in ('type modbus_analyzer.py ^| findstr /C:"VERSION"') do (
    set "line=%%a"
    set "line=!line:VERSION=!"
    set "line=!line: =!"
    set "line=!line:=!"
    set "line=!line:'=!"
    set "line=!line:"=!"
    set "version=!line!"
)

echo.
echo 创建发布包...

REM 创建release目录
mkdir release

REM 复制文件到release目录
echo 复制文件到release目录...
xcopy /E /I /Y "dist\modbus_analyzer\*" "release"

REM 创建README.txt
echo Modbus Analyzer v%version%> release\README.txt
echo.>> release\README.txt
echo 安装说明：>> release\README.txt
echo 1. 解压所有文件到任意目录>> release\README.txt
echo 2. 运行 modbus_analyzer.exe 启动程序>> release\README.txt
echo.>> release\README.txt
echo 注意事项：>> release\README.txt
echo - 首次运行可能需要管理员权限>> release\README.txt
echo - 如果运行时提示缺少DLL，请确保已安装Visual C++ Redistributable>> release\README.txt
echo - 配置文件位于程序目录下的config_and_params.xlsx>> release\README.txt
echo.>> release\README.txt
echo 如有问题，请联系技术支持。>> release\README.txt

REM 创建zip文件
set "zipFile=ModbusAnalyzer_v%version%.zip"
if exist "%zipFile%" del /f /q "%zipFile%"

echo 创建zip文件...
powershell -Command "Compress-Archive -Path 'release\*' -DestinationPath '%zipFile%'"

echo.
echo 打包和发布完成！
echo 发布包: %zipFile%
echo 发布目录: release
echo.
echo 你可以分发zip文件或release目录中的内容。
pause 