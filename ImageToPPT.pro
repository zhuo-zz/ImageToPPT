QT += widgets

CONFIG += c++17
QMAKE_CXXFLAGS += -finput-charset=UTF-8 -fexec-charset=UTF-8

TARGET = ImageToPPT
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
