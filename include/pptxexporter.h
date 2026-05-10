#pragma once

#include "Region.h"

#include <QImage>
#include <QString>
#include <QVector>

class PptxExporter {
public:
    bool exportPptx(const QString &outputPath,
                    const QImage &sourceImage,
                    const QVector<RegionItem> &regions,
                    bool includeOriginalBackground = true,
                    QString *errorMessage = nullptr);
    QImage buildPreviewImage(const QImage &sourceImage,
                             const QVector<RegionItem> &regions,
                             bool includeOriginalBackground = true) const;
    static bool isSimpleBackgroundRegion(const QImage &sourceImage, const QRect &rect);
    static QColor sampleRegionColor(const QImage &sourceImage, const QRect &rect);

private:
    static qint64 pxToEmu(int px, int imagePx, qint64 slideEmu);
    static QString xmlEscape(const QString &text);
    static QString colorToHex(const QColor &color);
    static bool writeTextFile(const QString &path, const QString &content, QString *errorMessage);
    static bool copyImageRegion(const QString &path, const QImage &sourceImage, const QRect &rect, QString *errorMessage);
    static QImage makeReconstructedBackground(const QImage &sourceImage, const QVector<RegionItem> &regions);
    static QRect backgroundRemovalRectForImageRegion(const QImage &sourceImage, const QRect &rect);
    static QVector<QColor> sampleStripColors(const QImage &image, const QRect &strip);
    static QColor medianColor(QVector<QColor> colors);
    static QColor medianColor(QVector<QColor> colors, const QColor &fallback);
    static int colorDistance(const QColor &a, const QColor &b);
    static QColor mixColor(const QColor &a, const QColor &b, double t);
    static bool isConsistentColorSet(const QVector<QColor> &colors, const QColor &center, int maxAverageDistance, double minNearRatio);
    static void fillRemovedImageRegion(QImage &image, const QRect &rect);
    static void fillRemovedTextRegion(QImage &image, const QRect &rect);

    QString buildContentTypes(int imageCount) const;
    QString buildRootRels() const;
    QString buildAppXml() const;
    QString buildCoreXml() const;
    QString buildPresentationXml(qint64 slideWidth, qint64 slideHeight) const;
    QString buildPresentationRels() const;
    QString buildSlideMasterXml() const;
    QString buildSlideMasterRels() const;
    QString buildSlideLayoutXml() const;
    QString buildSlideLayoutRels() const;
    QString buildSlideXml(const QImage &sourceImage,
                          const QVector<RegionItem> &regions,
                          qint64 slideWidth,
                          qint64 slideHeight,
                          bool includeOriginalBackground) const;
    QString buildSlideRels(int imageCount, bool includeOriginalBackground) const;
    QString buildThemeXml() const;
};
