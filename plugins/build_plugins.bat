@echo off
REM 编译插件为DLL文件

REM 创建build目录
if not exist "build" mkdir build
cd build

REM 配置CMake
cmake -G "Visual Studio 17 2022" -A x64 ..

REM 编译
cmake --build . --config Release

REM 返回上级目录
cd ..

echo.
echo 插件编译完成！请查看 dist 目录。
pause 