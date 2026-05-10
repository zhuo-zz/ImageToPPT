cmake_minimum_required(VERSION 3.16)

project(ImageToPPT LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTORCC ON)

find_package(QT NAMES Qt6 Qt5 REQUIRED COMPONENTS Widgets)
find_package(Qt${QT_VERSION_MAJOR} REQUIRED COMPONENTS Widgets)

add_executable(ImageToPPT
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

target_include_directories(ImageToPPT PRIVATE include)

if (CMAKE_CXX_COMPILER_ID MATCHES "GNU|Clang")
    target_compile_options(ImageToPPT PRIVATE -finput-charset=UTF-8 -fexec-charset=UTF-8)
endif()

target_link_libraries(ImageToPPT PRIVATE Qt${QT_VERSION_MAJOR}::Widgets)
