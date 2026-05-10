#pragma once

#include <QImage>
#include <QString>

struct PptxExportOptions
{
    int rows = 3;
    int columns = 3;
    bool addSmallGaps = false;
};

class PptxExporter
{
public:
    static bool exportImageGrid(const QImage &image,
                                const QString &outputPath,
                                const PptxExportOptions &options,
                                QString *errorMessage);
};
