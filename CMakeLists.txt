cmake_minimum_required(VERSION 3.16)

project(Image2Ppt LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTORCC ON)

find_package(QT NAMES Qt6 Qt5 REQUIRED COMPONENTS Widgets)
find_package(Qt${QT_VERSION_MAJOR} REQUIRED COMPONENTS Widgets)

add_executable(Image2Ppt
    source/main.cpp
    source/MainWindow.cpp
    include/MainWindow.h
    source/ImageCanvas.cpp
    include/ImageCanvas.h
    source/RegionDialog.cpp
    include/RegionDialog.h
    source/PptxExporter.cpp
    include/PptxExporter.h
    include/Region.h
)

target_include_directories(Image2Ppt PRIVATE include)

target_link_libraries(Image2Ppt PRIVATE Qt${QT_VERSION_MAJOR}::Widgets)
