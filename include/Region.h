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
    QString fontFamily = QStringLiteral("Microsoft YaHei");
    int fontSize = 24;
    QColor textColor = Qt::black;
    bool bold = false;
    bool hasImageBackgroundColor = false;
    QColor imageBackgroundColor = Qt::white;
};
