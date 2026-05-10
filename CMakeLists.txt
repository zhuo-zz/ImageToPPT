cmake_minimum_required(VERSION 3.16)

project(ImageToEditablePptx LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_AUTOMOC ON)

find_package(Qt6 COMPONENTS Widgets Gui QUIET)
if (Qt6_FOUND)
    set(QT_PACKAGE Qt6)
else()
    find_package(Qt5 COMPONENTS Widgets Gui REQUIRED)
    set(QT_PACKAGE Qt5)
endif()

add_executable(ImageToEditablePptx
    source/main.cpp
    source/mainwindow.cpp
    include/mainwindow.h
    source/pptxexporter.cpp
    include/pptxexporter.h
    source/zipwriter.cpp
    include/zipwriter.h
)

target_include_directories(ImageToEditablePptx PRIVATE include)

target_link_libraries(ImageToEditablePptx PRIVATE ${QT_PACKAGE}::Widgets ${QT_PACKAGE}::Gui)
