QT += widgets gui

CONFIG += c++17

TARGET = ImageToEditablePptx
TEMPLATE = app

SOURCES += \
    source/main.cpp \
    source/mainwindow.cpp \
    source/pptxexporter.cpp \
    source/zipwriter.cpp

HEADERS += \
    include/mainwindow.h \
    include/pptxexporter.h \
    include/zipwriter.h

INCLUDEPATH += include
