QT += widgets

CONFIG += c++17

TARGET = Image2Ppt
TEMPLATE = app

SOURCES += \
    source/main.cpp \
    source/MainWindow.cpp \
    source/ImageCanvas.cpp \
    source/RegionDialog.cpp \
    source/PptxExporter.cpp

HEADERS += \
    include/MainWindow.h \
    include/ImageCanvas.h \
    include/RegionDialog.h \
    include/PptxExporter.h \
    include/Region.h

INCLUDEPATH += include
