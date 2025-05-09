cmake_minimum_required(VERSION 3.4)
project(updater_plugin)

# 设置C++标准
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# 查找Python和pybind11
find_package(Python COMPONENTS Interpreter Development REQUIRED)
find_package(pybind11 CONFIG REQUIRED)

# 添加插件库
pybind11_add_module(updater_plugin updater_plugin.cpp)

# 设置输出目录
set_target_properties(updater_plugin PROPERTIES
    LIBRARY_OUTPUT_DIRECTORY "${CMAKE_SOURCE_DIR}/../dist"
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_SOURCE_DIR}/../dist"
) 