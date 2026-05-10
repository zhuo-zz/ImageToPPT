#include "PptxExporter.h"

#include <algorithm>

#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QFont>
#include <QPainter>
#include <QProcess>
#include <QTemporaryDir>
#include <QTextStream>
#include <QTextOption>
#include <QtGlobal>

#if QT_VERSION >= QT_VERSION_CHECK(6, 0, 0)
#include <QStringConverter>
#endif

#if QT_VERSION < QT_VERSION_CHECK(6, 0, 0)
#include <QTextCodec>
#endif

bool PptxExporter::exportPptx(const QString &outputPath,
                              const QImage &sourceImage,
                              const QVector<RegionItem> &regions,
                              bool includeOriginalBackground,
                              QString *errorMessage)
{
    if (sourceImage.isNull()) {
        if (errorMessage)
            *errorMessage = QStringLiteral("源图片为空。");
        return false;
    }

    QTemporaryDir tempDir;
    if (!tempDir.isValid()) {
        if (errorMessage)
            *errorMessage = QStringLiteral("无法创建临时目录。");
        return false;
    }

    const QString root = tempDir.path();
    QDir dir(root);
    dir.mkpath("_rels");
    dir.mkpath("docProps");
    dir.mkpath("ppt/_rels");
    dir.mkpath("ppt/slides/_rels");
    dir.mkpath("ppt/slideMasters/_rels");
    dir.mkpath("ppt/slideLayouts/_rels");
    dir.mkpath("ppt/theme");
    dir.mkpath("ppt/media");

    if (includeOriginalBackground) {
        const QString backgroundPath = root + QStringLiteral("/ppt/media/background.png");
        const QImage cleanBackground = makeReconstructedBackground(sourceImage, regions);
        if (!cleanBackground.save(backgroundPath, "PNG")) {
            if (errorMessage)
                *errorMessage = QStringLiteral("保存背景图失败。");
            return false;
        }
    }

    int imageCount = 0;
    for (const RegionItem &region : regions) {
        if (region.type != RegionType::Image)
            continue;

        ++imageCount;
        const QString mediaPath = root + QStringLiteral("/ppt/media/image%1.png").arg(imageCount);
        if (!copyImageRegion(mediaPath, sourceImage, region.imageRect, errorMessage))
            return false;
    }

    constexpr qint64 slideWidth = 12192000;  // 13.333 inch
    const qint64 slideHeight = qRound64(slideWidth * (static_cast<double>(sourceImage.height()) / sourceImage.width()));

    if (!writeTextFile(root + "/[Content_Types].xml", buildContentTypes(imageCount), errorMessage))
        return false;
    if (!writeTextFile(root + "/_rels/.rels", buildRootRels(), errorMessage))
        return false;
    if (!writeTextFile(root + "/docProps/app.xml", buildAppXml(), errorMessage))
        return false;
    if (!writeTextFile(root + "/docProps/core.xml", buildCoreXml(), errorMessage))
        return false;
    if (!writeTextFile(root + "/ppt/presentation.xml", buildPresentationXml(slideWidth, slideHeight), errorMessage))
        return false;
    if (!writeTextFile(root + "/ppt/_rels/presentation.xml.rels", buildPresentationRels(), errorMessage))
        return false;
    if (!writeTextFile(root + "/ppt/slideMasters/slideMaster1.xml", buildSlideMasterXml(), errorMessage))
        return false;
    if (!writeTextFile(root + "/ppt/slideMasters/_rels/slideMaster1.xml.rels", buildSlideMasterRels(), errorMessage))
        return false;
    if (!writeTextFile(root + "/ppt/slideLayouts/slideLayout1.xml", buildSlideLayoutXml(), errorMessage))
        return false;
    if (!writeTextFile(root + "/ppt/slideLayouts/_rels/slideLayout1.xml.rels", buildSlideLayoutRels(), errorMessage))
        return false;
    if (!writeTextFile(root + "/ppt/slides/slide1.xml", buildSlideXml(sourceImage, regions, slideWidth, slideHeight, includeOriginalBackground), errorMessage))
        return false;
    if (!writeTextFile(root + "/ppt/slides/_rels/slide1.xml.rels", buildSlideRels(imageCount, includeOriginalBackground), errorMessage))
        return false;
    if (!writeTextFile(root + "/ppt/theme/theme1.xml", buildThemeXml(), errorMessage))
        return false;

    const QString zipPath = outputPath + QStringLiteral(".zip");
    const QFileInfo outputInfo(outputPath);
    if (!outputInfo.absoluteDir().exists() && !QDir().mkpath(outputInfo.absolutePath())) {
        if (errorMessage)
            *errorMessage = QStringLiteral("无法创建输出目录：\n%1").arg(outputInfo.absolutePath());
        return false;
    }

    if (QFile::exists(outputPath) && !QFile::remove(outputPath)) {
        if (errorMessage)
            *errorMessage = QStringLiteral("无法覆盖输出文件：\n%1\n\n请先关闭正在打开这个 PPTX 的 PowerPoint/WPS/预览窗口，或换一个文件名再导出。").arg(outputPath);
        return false;
    }

    if (QFile::exists(zipPath) && !QFile::remove(zipPath)) {
        if (errorMessage)
            *errorMessage = QStringLiteral("无法删除旧的临时 ZIP：\n%1\n\n请检查文件是否被占用，或换一个文件名再导出。").arg(zipPath);
        return false;
    }

    const QString nativeRoot = QDir::toNativeSeparators(root).replace("'", "''");
    const QString nativeZipPath = QDir::toNativeSeparators(zipPath).replace("'", "''");
    const QString command = QStringLiteral(
        "$ErrorActionPreference='Stop';"
        "Add-Type -AssemblyName System.IO.Compression;"
        "Add-Type -AssemblyName System.IO.Compression.FileSystem;"
        "$root='%1';"
        "$zipPath='%2';"
        "if(Test-Path -LiteralPath $zipPath){Remove-Item -LiteralPath $zipPath -Force;}"
        "$fs=[System.IO.File]::Open($zipPath,[System.IO.FileMode]::CreateNew);"
        "$zip=New-Object System.IO.Compression.ZipArchive($fs,[System.IO.Compression.ZipArchiveMode]::Create);"
        "try{"
        "Get-ChildItem -LiteralPath $root -Recurse -File | ForEach-Object {"
        "$entryName=$_.FullName.Substring($root.Length+1).Replace('\\','/');"
        "[System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip,$_.FullName,$entryName,[System.IO.Compression.CompressionLevel]::Optimal) | Out-Null;"
        "}"
        "}finally{$zip.Dispose();$fs.Dispose();}")
        .arg(nativeRoot, nativeZipPath);

    QProcess process;
    process.start("powershell", {"-NoProfile", "-Command", command});
    if (!process.waitForFinished(30000) || process.exitCode() != 0) {
        if (errorMessage) {
            *errorMessage = QStringLiteral("PPTX 打包失败：\n%1")
                .arg(QString::fromLocal8Bit(process.readAllStandardError()));
        }
        return false;
    }

    if (!QFile::rename(zipPath, outputPath)) {
        if (errorMessage)
            *errorMessage = QStringLiteral("无法将临时 ZIP 重命名为 PPTX：\n%1").arg(outputPath);
        QFile::remove(zipPath);
        return false;
    }

    return true;
}

QImage PptxExporter::buildPreviewImage(const QImage &sourceImage,
                                       const QVector<RegionItem> &regions,
                                       bool includeOriginalBackground) const
{
    if (sourceImage.isNull())
        return {};

    QImage preview = includeOriginalBackground
        ? makeReconstructedBackground(sourceImage, regions)
        : QImage(sourceImage.size(), QImage::Format_ARGB32);
    if (!includeOriginalBackground)
        preview.fill(Qt::white);

    QPainter painter(&preview);
    painter.setRenderHint(QPainter::Antialiasing);
    painter.setRenderHint(QPainter::TextAntialiasing);

    for (const RegionItem &region : regions) {
        const QRect rect = region.imageRect.intersected(sourceImage.rect());
        if (rect.isEmpty())
            continue;

        if (region.type == RegionType::Image) {
            painter.drawImage(rect, sourceImage.copy(rect));
            continue;
        }

        QFont font(region.fontFamily.isEmpty() ? QStringLiteral("Microsoft YaHei") : region.fontFamily);
        const int pixelSize = qBound(8, qRound(region.fontSize * sourceImage.width() / 960.0), 160);
        font.setPixelSize(pixelSize);
        font.setBold(region.bold);
        painter.setFont(font);
        painter.setPen(region.textColor);

        QTextOption option;
        option.setWrapMode(QTextOption::WrapAtWordBoundaryOrAnywhere);
        option.setAlignment(Qt::AlignLeft | Qt::AlignTop);
        painter.drawText(QRectF(rect), region.text, option);
    }

    return preview;
}

bool PptxExporter::isSimpleBackgroundRegion(const QImage &sourceImage, const QRect &rect)
{
    const QRect bounded = rect.intersected(sourceImage.rect());
    if (sourceImage.isNull() || bounded.isEmpty())
        return false;

    QVector<QColor> samples;
    samples.reserve(bounded.width() * bounded.height());
    const int step = qMax(1, qMin(bounded.width(), bounded.height()) / 80);
    for (int y = bounded.top(); y <= bounded.bottom(); y += step) {
        for (int x = bounded.left(); x <= bounded.right(); x += step)
            samples.append(sourceImage.pixelColor(x, y));
    }

    const QColor center = medianColor(samples);
    int nearCount = 0;
    int distanceSum = 0;
    for (const QColor &color : samples) {
        const int distance = colorDistance(color, center);
        distanceSum += qMin(distance, 255);
        if (distance <= 54)
            ++nearCount;
    }

    if (samples.isEmpty())
        return false;

    const double nearRatio = static_cast<double>(nearCount) / samples.size();
    const double averageDistance = static_cast<double>(distanceSum) / samples.size();
    return nearRatio >= 0.72 && averageDistance <= 48.0;
}

QColor PptxExporter::sampleRegionColor(const QImage &sourceImage, const QRect &rect)
{
    const QRect bounded = rect.intersected(sourceImage.rect());
    if (sourceImage.isNull() || bounded.isEmpty())
        return Qt::white;

    QVector<QColor> colors;
    const int step = qMax(1, qMin(bounded.width(), bounded.height()) / 60);
    for (int y = bounded.top(); y <= bounded.bottom(); y += step) {
        for (int x = bounded.left(); x <= bounded.right(); x += step)
            colors.append(sourceImage.pixelColor(x, y));
    }

    return medianColor(colors);
}

qint64 PptxExporter::pxToEmu(int px, int imagePx, qint64 slideEmu)
{
    return qRound64(static_cast<double>(px) / imagePx * slideEmu);
}

QString PptxExporter::xmlEscape(const QString &text)
{
    QString escaped = text;
    escaped.replace("&", "&amp;");
    escaped.replace("<", "&lt;");
    escaped.replace(">", "&gt;");
    escaped.replace("\"", "&quot;");
    escaped.replace("'", "&apos;");
    return escaped;
}

QString PptxExporter::colorToHex(const QColor &color)
{
    return QStringLiteral("%1%2%3")
        .arg(color.red(), 2, 16, QLatin1Char('0'))
        .arg(color.green(), 2, 16, QLatin1Char('0'))
        .arg(color.blue(), 2, 16, QLatin1Char('0'))
        .toUpper();
}

bool PptxExporter::writeTextFile(const QString &path, const QString &content, QString *errorMessage)
{
    QFile file(path);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Text)) {
        if (errorMessage)
            *errorMessage = QStringLiteral("无法写入文件：%1").arg(path);
        return false;
    }

    QTextStream out(&file);
#if QT_VERSION >= QT_VERSION_CHECK(6, 0, 0)
    out.setEncoding(QStringConverter::Utf8);
#else
    out.setCodec(QTextCodec::codecForName("UTF-8"));
#endif
    out << content;
    return true;
}

bool PptxExporter::copyImageRegion(const QString &path, const QImage &sourceImage, const QRect &rect, QString *errorMessage)
{
    const QRect bounded = rect.intersected(sourceImage.rect());
    if (bounded.isEmpty()) {
        if (errorMessage)
            *errorMessage = QStringLiteral("图片区域为空。");
        return false;
    }

    if (!sourceImage.copy(bounded).save(path, "PNG")) {
        if (errorMessage)
            *errorMessage = QStringLiteral("图片裁剪保存失败：%1").arg(path);
        return false;
    }

    return true;
}

QImage PptxExporter::makeReconstructedBackground(const QImage &sourceImage, const QVector<RegionItem> &regions)
{
    QImage background = sourceImage.convertToFormat(QImage::Format_ARGB32);

    for (const RegionItem &region : regions) {
        QRect rect = region.imageRect.intersected(background.rect());
        if (rect.isEmpty())
            continue;

        if (region.type == RegionType::Image) {
            const QRect removalRect = backgroundRemovalRectForImageRegion(sourceImage, rect);
            if (region.hasImageBackgroundColor) {
                QPainter painter(&background);
                painter.fillRect(removalRect, region.imageBackgroundColor);
            } else {
                fillRemovedImageRegion(background, removalRect);
            }
        } else {
            fillRemovedTextRegion(background, rect);
        }
    }

    return background;
}

QRect PptxExporter::backgroundRemovalRectForImageRegion(const QImage &sourceImage, const QRect &rect)
{
    const QRect bounded = rect.intersected(sourceImage.rect());
    if (sourceImage.isNull() || bounded.isEmpty())
        return {};

    const int bleed = 2;
    return bounded.adjusted(-bleed, -bleed, bleed, bleed).intersected(sourceImage.rect());
}

QVector<QColor> PptxExporter::sampleStripColors(const QImage &image, const QRect &strip)
{
    QVector<QColor> colors;
    const QRect bounded = strip.intersected(image.rect());
    colors.reserve(bounded.width() * bounded.height());

    for (int y = bounded.top(); y <= bounded.bottom(); ++y) {
        for (int x = bounded.left(); x <= bounded.right(); ++x) {
            colors.append(image.pixelColor(x, y));
        }
    }

    return colors;
}

QColor PptxExporter::medianColor(QVector<QColor> colors)
{
    if (colors.isEmpty())
        return Qt::white;

    QVector<int> reds;
    QVector<int> greens;
    QVector<int> blues;
    reds.reserve(colors.size());
    greens.reserve(colors.size());
    blues.reserve(colors.size());

    for (const QColor &color : colors) {
        reds.append(color.red());
        greens.append(color.green());
        blues.append(color.blue());
    }

    auto median = [](QVector<int> values) {
        std::sort(values.begin(), values.end());
        return values.at(values.size() / 2);
    };

    return QColor(median(reds), median(greens), median(blues));
}

QColor PptxExporter::medianColor(QVector<QColor> colors, const QColor &fallback)
{
    if (colors.isEmpty())
        return fallback;

    return medianColor(colors);
}

int PptxExporter::colorDistance(const QColor &a, const QColor &b)
{
    return qAbs(a.red() - b.red())
        + qAbs(a.green() - b.green())
        + qAbs(a.blue() - b.blue());
}

QColor PptxExporter::mixColor(const QColor &a, const QColor &b, double t)
{
    t = qBound(0.0, t, 1.0);
    return QColor(
        qRound(a.red() * (1.0 - t) + b.red() * t),
        qRound(a.green() * (1.0 - t) + b.green() * t),
        qRound(a.blue() * (1.0 - t) + b.blue() * t)
    );
}

bool PptxExporter::isConsistentColorSet(const QVector<QColor> &colors, const QColor &center, int maxAverageDistance, double minNearRatio)
{
    if (colors.isEmpty())
        return false;

    int nearCount = 0;
    int distanceSum = 0;
    for (const QColor &color : colors) {
        const int distance = colorDistance(color, center);
        distanceSum += qMin(distance, 255);
        if (distance <= maxAverageDistance * 2)
            ++nearCount;
    }

    const double nearRatio = static_cast<double>(nearCount) / colors.size();
    const double averageDistance = static_cast<double>(distanceSum) / colors.size();
    return nearRatio >= minNearRatio && averageDistance <= maxAverageDistance;
}

void PptxExporter::fillRemovedImageRegion(QImage &image, const QRect &rect)
{
    const QRect bounded = rect.intersected(image.rect());
    if (bounded.isEmpty())
        return;

    const QImage original = image.copy();
    const int margin = qBound(12, qMin(bounded.width(), bounded.height()) / 5, 42);
    const int rowRadius = qBound(2, bounded.height() / 24, 8);
    const int columnRadius = qBound(2, bounded.width() / 24, 8);

    QVector<QColor> ringColors;
    ringColors += sampleStripColors(original, QRect(bounded.left() - margin, bounded.top(), margin, bounded.height()));
    ringColors += sampleStripColors(original, QRect(bounded.right() + 1, bounded.top(), margin, bounded.height()));
    ringColors += sampleStripColors(original, QRect(bounded.left(), bounded.top() - margin, bounded.width(), margin));
    ringColors += sampleStripColors(original, QRect(bounded.left(), bounded.bottom() + 1, bounded.width(), margin));
    const QColor fallback = medianColor(ringColors, medianColor(sampleStripColors(original, bounded)));
    if (isConsistentColorSet(ringColors, fallback, 18, 0.82)) {
        QPainter painter(&image);
        painter.fillRect(bounded, fallback);
        return;
    }

    QVector<QColor> topColors;
    QVector<QColor> bottomColors;
    topColors.reserve(bounded.width());
    bottomColors.reserve(bounded.width());
    for (int x = bounded.left(); x <= bounded.right(); ++x) {
        topColors.append(medianColor(
            sampleStripColors(original, QRect(x - columnRadius, bounded.top() - margin, columnRadius * 2 + 1, margin)),
            fallback
        ));
        bottomColors.append(medianColor(
            sampleStripColors(original, QRect(x - columnRadius, bounded.bottom() + 1, columnRadius * 2 + 1, margin)),
            fallback
        ));
    }

    for (int y = bounded.top(); y <= bounded.bottom(); ++y) {
        const double ty = bounded.height() <= 1
            ? 0.0
            : static_cast<double>(y - bounded.top()) / (bounded.height() - 1);

        const QColor leftColor = medianColor(
            sampleStripColors(original, QRect(bounded.left() - margin, y - rowRadius, margin, rowRadius * 2 + 1)),
            fallback
        );
        const QColor rightColor = medianColor(
            sampleStripColors(original, QRect(bounded.right() + 1, y - rowRadius, margin, rowRadius * 2 + 1)),
            fallback
        );

        for (int x = bounded.left(); x <= bounded.right(); ++x) {
            const double tx = bounded.width() <= 1
                ? 0.0
                : static_cast<double>(x - bounded.left()) / (bounded.width() - 1);

            const QColor horizontal = mixColor(leftColor, rightColor, tx);
            const QColor vertical = mixColor(topColors.at(x - bounded.left()), bottomColors.at(x - bounded.left()), ty);
            const QColor repaired = mixColor(horizontal, vertical, 0.52);
            image.setPixelColor(x, y, repaired);
        }
    }
}

void PptxExporter::fillRemovedTextRegion(QImage &image, const QRect &rect)
{
    const QRect bounded = rect.intersected(image.rect());
    if (bounded.isEmpty())
        return;

    const QImage original = image.copy();
    const int margin = qBound(8, qMin(bounded.width(), bounded.height()) / 5, 32);
    const int rowRadius = qBound(1, bounded.height() / 28, 4);
    const int columnRadius = qBound(1, bounded.width() / 28, 4);
    const int localRadius = qBound(2, qMin(bounded.width(), bounded.height()) / 16, 6);

    QVector<QColor> ringColors;
    ringColors += sampleStripColors(original, QRect(bounded.left() - margin, bounded.top(), margin, bounded.height()));
    ringColors += sampleStripColors(original, QRect(bounded.right() + 1, bounded.top(), margin, bounded.height()));
    ringColors += sampleStripColors(original, QRect(bounded.left(), bounded.top() - margin, bounded.width(), margin));
    ringColors += sampleStripColors(original, QRect(bounded.left(), bounded.bottom() + 1, bounded.width(), margin));
    const QColor fallback = medianColor(ringColors, medianColor(sampleStripColors(original, bounded)));
    int ringNearCount = 0;
    int ringDistanceSum = 0;
    for (const QColor &color : ringColors) {
        const int distance = colorDistance(color, fallback);
        ringDistanceSum += distance;
        if (distance <= 48)
            ++ringNearCount;
    }
    const bool simpleSurroundingBackground = ringColors.size() >= qMax(16, bounded.width() + bounded.height())
        && static_cast<double>(ringNearCount) / ringColors.size() >= 0.78
        && static_cast<double>(ringDistanceSum) / ringColors.size() <= 34.0;

    QVector<QColor> estimatedFill;
    estimatedFill.reserve(bounded.width() * bounded.height());

    QVector<QColor> topColors;
    QVector<QColor> bottomColors;
    topColors.reserve(bounded.width());
    bottomColors.reserve(bounded.width());
    for (int x = bounded.left(); x <= bounded.right(); ++x) {
        topColors.append(medianColor(
            sampleStripColors(original, QRect(x - columnRadius, bounded.top() - margin, columnRadius * 2 + 1, margin)),
            fallback
        ));
        bottomColors.append(medianColor(
            sampleStripColors(original, QRect(x - columnRadius, bounded.bottom() + 1, columnRadius * 2 + 1, margin)),
            fallback
        ));
    }

    for (int y = bounded.top(); y <= bounded.bottom(); ++y) {
        const double ty = bounded.height() <= 1
            ? 0.0
            : static_cast<double>(y - bounded.top()) / (bounded.height() - 1);

        const QColor leftColor = medianColor(
            sampleStripColors(original, QRect(bounded.left() - margin, y - rowRadius, margin, rowRadius * 2 + 1)),
            fallback
        );
        const QColor rightColor = medianColor(
            sampleStripColors(original, QRect(bounded.right() + 1, y - rowRadius, margin, rowRadius * 2 + 1)),
            fallback
        );

        for (int x = bounded.left(); x <= bounded.right(); ++x) {
            const double tx = bounded.width() <= 1
                ? 0.0
                : static_cast<double>(x - bounded.left()) / (bounded.width() - 1);

            QColor horizontal = mixColor(leftColor, rightColor, tx);
            QColor vertical = mixColor(topColors.at(x - bounded.left()), bottomColors.at(x - bounded.left()), ty);
            estimatedFill.append(mixColor(horizontal, vertical, 0.5));
        }
    }

    const int width = bounded.width();
    const int height = bounded.height();
    QVector<uchar> textMask(width * height, 0);
    QVector<int> luminances;
    luminances.reserve(width * height);
    for (int y = bounded.top(); y <= bounded.bottom(); ++y) {
        for (int x = bounded.left(); x <= bounded.right(); ++x) {
            const QColor pixel = original.pixelColor(x, y);
            luminances.append(qRound(pixel.red() * 0.299 + pixel.green() * 0.587 + pixel.blue() * 0.114));
        }
    }
    std::sort(luminances.begin(), luminances.end());

    auto luminance = [](const QColor &color) {
        return qRound(color.red() * 0.299 + color.green() * 0.587 + color.blue() * 0.114);
    };

    const int lowLum = luminances.at(qBound(0, luminances.size() / 10, luminances.size() - 1));
    const int highLum = luminances.at(qBound(0, luminances.size() * 9 / 10, luminances.size() - 1));
    const int backgroundLum = luminance(fallback);
    const bool hasDarkText = lowLum < backgroundLum - 22;
    const bool hasLightText = highLum > backgroundLum + 28;
    const int strongThreshold = 118;
    const int weakThreshold = 58;

    QVector<int> dominantBuckets(4096, 0);
    for (int y = bounded.top(); y <= bounded.bottom(); ++y) {
        for (int x = bounded.left(); x <= bounded.right(); ++x) {
            const QColor pixel = original.pixelColor(x, y);
            const int key = ((pixel.red() >> 4) << 8) | ((pixel.green() >> 4) << 4) | (pixel.blue() >> 4);
            ++dominantBuckets[key];
        }
    }

    int dominantKey = 0;
    int dominantCount = 0;
    for (int i = 0; i < dominantBuckets.size(); ++i) {
        if (dominantBuckets.at(i) > dominantCount) {
            dominantCount = dominantBuckets.at(i);
            dominantKey = i;
        }
    }

    QVector<QColor> dominantColors;
    dominantColors.reserve(dominantCount);
    for (int y = bounded.top(); y <= bounded.bottom(); ++y) {
        for (int x = bounded.left(); x <= bounded.right(); ++x) {
            const QColor pixel = original.pixelColor(x, y);
            const int key = ((pixel.red() >> 4) << 8) | ((pixel.green() >> 4) << 4) | (pixel.blue() >> 4);
            if (key == dominantKey)
                dominantColors.append(pixel);
        }
    }

    const QColor dominantBackground = medianColor(dominantColors, fallback);
    int dominantNearCount = 0;
    int dominantDistanceSum = 0;
    for (int y = bounded.top(); y <= bounded.bottom(); ++y) {
        for (int x = bounded.left(); x <= bounded.right(); ++x) {
            const int distance = colorDistance(original.pixelColor(x, y), dominantBackground);
            dominantDistanceSum += qMin(distance, 180);
            if (distance <= 54)
                ++dominantNearCount;
        }
    }
    const int pixelCount = width * height;
    const double dominantRatio = pixelCount == 0 ? 0.0 : static_cast<double>(dominantCount) / pixelCount;
    const double dominantNearRatio = pixelCount == 0 ? 0.0 : static_cast<double>(dominantNearCount) / pixelCount;
    const double dominantAverageDistance = pixelCount == 0 ? 999.0 : static_cast<double>(dominantDistanceSum) / pixelCount;
    const bool dominantColorBackground = dominantRatio >= 0.18
        && dominantNearRatio >= 0.48
        && dominantAverageDistance <= 82.0;

    for (int y = bounded.top(); y <= bounded.bottom(); ++y) {
        for (int x = bounded.left(); x <= bounded.right(); ++x) {
            const int index = (y - bounded.top()) * width + (x - bounded.left());
            const QColor pixel = original.pixelColor(x, y);
            const QColor localBackground = medianColor(
                sampleStripColors(original, QRect(x - localRadius, y - localRadius, localRadius * 2 + 1, localRadius * 2 + 1)),
                estimatedFill.at(index)
            );

            const int localDistance = colorDistance(pixel, localBackground);
            const int fillDistance = colorDistance(pixel, estimatedFill.at(index));
            const int lum = luminance(pixel);
            const int localLum = luminance(localBackground);
            const int fillLum = luminance(estimatedFill.at(index));
            const bool darkForeground = hasDarkText
                && (lum <= localLum - 18 || lum <= fillLum - 20 || lum <= qMin(lowLum + 24, backgroundLum - 18));
            const bool lightForeground = hasLightText
                && (lum >= localLum + 22 || lum >= fillLum + 24 || lum >= qMax(highLum - 24, backgroundLum + 22));
            const bool strongForeground = (darkForeground || lightForeground)
                && qMax(localDistance, fillDistance) >= strongThreshold;
            const bool weakSmallTextForeground = (darkForeground || lightForeground)
                && qMax(localDistance, fillDistance) >= weakThreshold
                && qAbs(lum - localLum) >= 16;
            const bool largeGlyphForeground = (simpleSurroundingBackground || dominantColorBackground)
                && qMax(fillDistance, qMax(colorDistance(pixel, fallback), colorDistance(pixel, dominantBackground))) >= 74
                && qAbs(lum - backgroundLum) >= 16;

            if (strongForeground || weakSmallTextForeground || largeGlyphForeground)
                textMask[index] = 1;
        }
    }

    QVector<uchar> visited(width * height, 0);
    const int maxTextComponentArea = qMax(24, width * height / 14);
    for (int start = 0; start < textMask.size(); ++start) {
        if (!textMask.at(start) || visited.at(start))
            continue;

        QVector<int> stack;
        QVector<int> component;
        stack.append(start);
        visited[start] = 1;

        int minX = width;
        int minY = height;
        int maxX = 0;
        int maxY = 0;

        while (!stack.isEmpty()) {
            const int current = stack.takeLast();
            component.append(current);
            const int cx = current % width;
            const int cy = current / width;
            minX = qMin(minX, cx);
            minY = qMin(minY, cy);
            maxX = qMax(maxX, cx);
            maxY = qMax(maxY, cy);

            for (int dy = -1; dy <= 1; ++dy) {
                for (int dx = -1; dx <= 1; ++dx) {
                    if (dx == 0 && dy == 0)
                        continue;

                    const int nx = cx + dx;
                    const int ny = cy + dy;
                    if (nx < 0 || nx >= width || ny < 0 || ny >= height)
                        continue;

                    const int next = ny * width + nx;
                    if (!textMask.at(next) || visited.at(next))
                        continue;

                    visited[next] = 1;
                    stack.append(next);
                }
            }
        }

        const int componentWidth = maxX - minX + 1;
        const int componentHeight = maxY - minY + 1;
        const bool tooLarge = !simpleSurroundingBackground && component.size() > maxTextComponentArea;
        const bool textureLike = !simpleSurroundingBackground && componentWidth > width * 0.72 && componentHeight > height * 0.32;
        if (tooLarge || textureLike) {
            for (int index : component)
                textMask[index] = 0;
        }
    }

    QVector<QColor> backgroundSamples;
    backgroundSamples.reserve(width * height);
    for (int y = bounded.top(); y <= bounded.bottom(); ++y) {
        for (int x = bounded.left(); x <= bounded.right(); ++x) {
            const int index = (y - bounded.top()) * width + (x - bounded.left());
            if (!textMask.at(index))
                backgroundSamples.append(original.pixelColor(x, y));
        }
    }

    const QColor flatBackground = medianColor(backgroundSamples, fallback);
    int flatNearCount = 0;
    int flatDistanceSum = 0;
    for (const QColor &color : backgroundSamples) {
        const int distance = colorDistance(color, flatBackground);
        flatDistanceSum += distance;
        if (distance <= 36)
            ++flatNearCount;
    }
    const double flatNearRatio = backgroundSamples.isEmpty()
        ? 0.0
        : static_cast<double>(flatNearCount) / backgroundSamples.size();
    const double flatAverageDistance = backgroundSamples.isEmpty()
        ? 999.0
        : static_cast<double>(flatDistanceSum) / backgroundSamples.size();
    const bool flatColorBackgroundBySamples = backgroundSamples.size() >= qMax(12, width * height / 8)
        && flatNearRatio >= 0.82
        && flatAverageDistance <= 24.0;
    const bool flatColorBackground = flatColorBackgroundBySamples || dominantColorBackground;
    const QColor repairBackground = flatColorBackgroundBySamples ? flatBackground : dominantBackground;

    if (flatColorBackground) {
        const int flatLum = luminance(repairBackground);
        for (int y = bounded.top(); y <= bounded.bottom(); ++y) {
            for (int x = bounded.left(); x <= bounded.right(); ++x) {
                const int index = (y - bounded.top()) * width + (x - bounded.left());
                const QColor pixel = original.pixelColor(x, y);
                const int distance = colorDistance(pixel, repairBackground);
                const int lum = luminance(pixel);
                const bool darkOnFlat = lum <= flatLum - 10;
                const bool lightOnFlat = lum >= flatLum + 10;
                const bool colorShiftOnFlat = distance >= 22 && qAbs(lum - flatLum) >= 8;
                if ((darkOnFlat || lightOnFlat || colorShiftOnFlat) && distance >= 22)
                    textMask[index] = 1;
            }
        }
    }

    QVector<uchar> eraseMask = textMask;
    const int dilateRadius = flatColorBackground
        ? qBound(1, qMin(width, height) / 30, 3)
        : qBound(1, qMin(width, height) / 36, 2);
    for (int y = 0; y < height; ++y) {
        for (int x = 0; x < width; ++x) {
            if (!textMask[y * width + x])
                continue;

            for (int yy = qMax(0, y - dilateRadius); yy <= qMin(height - 1, y + dilateRadius); ++yy) {
                for (int xx = qMax(0, x - dilateRadius); xx <= qMin(width - 1, x + dilateRadius); ++xx)
                    eraseMask[yy * width + xx] = 1;
            }
        }
    }

    for (int y = bounded.top(); y <= bounded.bottom(); ++y) {
        for (int x = bounded.left(); x <= bounded.right(); ++x) {
            const int index = (y - bounded.top()) * width + (x - bounded.left());
            if (!eraseMask.at(index))
                continue;

            if (flatColorBackground) {
                image.setPixelColor(x, y, repairBackground);
                continue;
            }

            const int localMargin = qBound(4, qMin(width, height) / 12, 10);
            QVector<QColor> nearby;
            const QRect around(x - localMargin, y - localMargin, localMargin * 2 + 1, localMargin * 2 + 1);
            const QRect clipped = around.intersected(bounded);
            for (int yy = clipped.top(); yy <= clipped.bottom(); ++yy) {
                for (int xx = clipped.left(); xx <= clipped.right(); ++xx) {
                    const int nearbyIndex = (yy - bounded.top()) * width + (xx - bounded.left());
                    if (!eraseMask.at(nearbyIndex))
                        nearby.append(original.pixelColor(xx, yy));
                }
            }

            QColor repaired = medianColor(nearby, estimatedFill.at(index));
            repaired = mixColor(repaired, estimatedFill.at(index), 0.25);
            image.setPixelColor(x, y, repaired);
        }
    }
}

QString PptxExporter::buildContentTypes(int imageCount) const
{
    QString xml = QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>
)");
    Q_UNUSED(imageCount);
    return xml;
}

QString PptxExporter::buildRootRels() const
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
)");
}

QString PptxExporter::buildAppXml() const
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>ImageToPPT</Application>
  <PresentationFormat>Custom</PresentationFormat>
  <Slides>1</Slides>
  <Notes>0</Notes>
  <HiddenSlides>0</HiddenSlides>
  <Company></Company>
  <AppVersion>1.0</AppVersion>
</Properties>
)");
}

QString PptxExporter::buildCoreXml() const
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>ImageToPPT Export</dc:title>
  <dc:creator>ImageToPPT</dc:creator>
  <cp:lastModifiedBy>ImageToPPT</cp:lastModifiedBy>
</cp:coreProperties>
)");
}

QString PptxExporter::buildPresentationXml(qint64 slideWidth, qint64 slideHeight) const
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId3"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
  <p:sldSz cx="%1" cy="%2"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>
)").arg(slideWidth).arg(slideHeight);
}

QString PptxExporter::buildPresentationRels() const
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
</Relationships>
)");
}

QString PptxExporter::buildSlideMasterXml() const
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle/>
    <p:bodyStyle/>
    <p:otherStyle/>
  </p:txStyles>
</p:sldMaster>
)");
}

QString PptxExporter::buildSlideMasterRels() const
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>
)");
}

QString PptxExporter::buildSlideLayoutXml() const
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>
)");
}

QString PptxExporter::buildSlideLayoutRels() const
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>
)");
}

QString PptxExporter::buildSlideXml(const QImage &sourceImage,
                                    const QVector<RegionItem> &regions,
                                    qint64 slideWidth,
                                    qint64 slideHeight,
                                    bool includeOriginalBackground) const
{
    QString shapes;
    int shapeId = 1;
    int imageRelId = includeOriginalBackground ? 2 : 1;

    if (includeOriginalBackground) {
        ++shapeId;
        shapes += QStringLiteral(R"(
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="%1" name="Original Background"/>
          <p:cNvPicPr/>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="rId2"/>
          <a:stretch><a:fillRect/></a:stretch>
        </p:blipFill>
        <p:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="%2" cy="%3"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
      </p:pic>)").arg(shapeId).arg(slideWidth).arg(slideHeight);
    }

    for (const RegionItem &region : regions) {
        const QRect r = region.imageRect.intersected(sourceImage.rect());
        const qint64 x = pxToEmu(r.x(), sourceImage.width(), slideWidth);
        const qint64 y = pxToEmu(r.y(), sourceImage.height(), slideHeight);
        const qint64 w = pxToEmu(r.width(), sourceImage.width(), slideWidth);
        const qint64 h = pxToEmu(r.height(), sourceImage.height(), slideHeight);

        if (region.type == RegionType::Image) {
            ++shapeId;
            ++imageRelId;
            shapes += QStringLiteral(R"(
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="%1" name="Picture %1"/>
          <p:cNvPicPr/>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="rId%2"/>
          <a:stretch><a:fillRect/></a:stretch>
        </p:blipFill>
        <p:spPr>
          <a:xfrm>
            <a:off x="%3" y="%4"/>
            <a:ext cx="%5" cy="%6"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
      </p:pic>)").arg(shapeId).arg(imageRelId).arg(x).arg(y).arg(w).arg(h);
        } else {
            ++shapeId;
            const int fontSize = qBound(8, region.fontSize, 96) * 100;
            const QString bold = region.bold ? QStringLiteral("1") : QStringLiteral("0");
            const QStringList lines = region.text.split('\n');
            QString paragraphs;

            for (const QString &line : lines) {
                paragraphs += QStringLiteral(R"(
          <a:p>
            <a:r>
              <a:rPr lang="zh-CN" sz="%1" b="%2">
                <a:solidFill><a:srgbClr val="%3"/></a:solidFill>
                <a:latin typeface="%5"/>
                <a:ea typeface="%5"/>
              </a:rPr>
              <a:t>%6</a:t>
            </a:r>
          </a:p>)").arg(fontSize)
                 .arg(bold)
                 .arg(colorToHex(region.textColor))
                 .arg(xmlEscape(region.fontFamily.isEmpty() ? QStringLiteral("Microsoft YaHei") : region.fontFamily))
                 .arg(xmlEscape(line));
            }

            shapes += QStringLiteral(R"(
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="%1" name="TextBox %1"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="%2" y="%3"/>
            <a:ext cx="%4" cy="%5"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
          <a:ln><a:noFill/></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0"/>
          <a:lstStyle/>
%6
        </p:txBody>
      </p:sp>)").arg(shapeId).arg(x).arg(y).arg(w).arg(h).arg(paragraphs);
        }
    }

    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>%1
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
)").arg(shapes);
}

QString PptxExporter::buildSlideRels(int imageCount, bool includeOriginalBackground) const
{
    QString rels = QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
)");

    int relId = 2;

    if (includeOriginalBackground) {
        rels += QStringLiteral("  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/background.png\"/>\n");
        relId = 3;
    }

    for (int i = 1; i <= imageCount; ++i) {
        rels += QStringLiteral("  <Relationship Id=\"rId%1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image%2.png\"/>\n")
            .arg(relId++)
            .arg(i);
    }

    rels += QStringLiteral("</Relationships>\n");
    return rels;
}

QString PptxExporter::buildThemeXml() const
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="ImageToPPT">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F1F1F"/></a:dk2>
      <a:lt2><a:srgbClr val="F2F2F2"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="ImageToPPT">
      <a:majorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/><a:cs typeface="Arial"/></a:majorFont>
      <a:minorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/><a:cs typeface="Arial"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="ImageToPPT">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>
)");
}
