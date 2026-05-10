#pragma once

#include <QColor>
#include <QRect>
#include <QString>

enum class RegionType {
    Image,
    Text
};

struct RegionItem {
    RegionType type = RegionType::Image;
    QRect imageRect;
    QString text;
    int fontSize = 24;
    QColor textColor = Qt::black;
    bool bold = false;
};
