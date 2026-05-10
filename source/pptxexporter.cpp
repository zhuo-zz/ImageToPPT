#include "PptxExporter.h"

#include <algorithm>

#include <QDir>
#include <QFile>
#include <QProcess>
#include <QTemporaryDir>
#include <QTextStream>
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
        const QImage cleanBackground = makeBackgroundWithoutText(sourceImage, regions);
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

    QFile::remove(outputPath);
    const QString zipPath = outputPath + QStringLiteral(".zip");
    QFile::remove(zipPath);

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

QImage PptxExporter::makeBackgroundWithoutText(const QImage &sourceImage, const QVector<RegionItem> &regions)
{
    QImage background = sourceImage.convertToFormat(QImage::Format_ARGB32);

    for (const RegionItem &region : regions) {
        if (region.type != RegionType::Text)
            continue;

        QRect rect = region.imageRect.intersected(background.rect());
        if (rect.isEmpty())
            continue;

        // Slightly enlarge the erase area so anti-aliased character edges do not remain.
        rect = rect.adjusted(-3, -3, 3, 3).intersected(background.rect());
        fillRemovedTextRegion(background, rect);
    }

    return background;
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

QColor PptxExporter::mixColor(const QColor &a, const QColor &b, double t)
{
    t = qBound(0.0, t, 1.0);
    return QColor(
        qRound(a.red() * (1.0 - t) + b.red() * t),
        qRound(a.green() * (1.0 - t) + b.green() * t),
        qRound(a.blue() * (1.0 - t) + b.blue() * t)
    );
}

void PptxExporter::fillRemovedTextRegion(QImage &image, const QRect &rect)
{
    const QRect bounded = rect.intersected(image.rect());
    if (bounded.isEmpty())
        return;

    const int margin = qBound(6, qMin(bounded.width(), bounded.height()) / 8, 24);

    const QColor leftColor = medianColor(sampleStripColors(
        image,
        QRect(bounded.left() - margin, bounded.top(), margin, bounded.height())
    ));
    const QColor rightColor = medianColor(sampleStripColors(
        image,
        QRect(bounded.right() + 1, bounded.top(), margin, bounded.height())
    ));
    const QColor topColor = medianColor(sampleStripColors(
        image,
        QRect(bounded.left(), bounded.top() - margin, bounded.width(), margin)
    ));
    const QColor bottomColor = medianColor(sampleStripColors(
        image,
        QRect(bounded.left(), bounded.bottom() + 1, bounded.width(), margin)
    ));

    QColor fallback = medianColor(sampleStripColors(image, bounded.adjusted(-margin, -margin, margin, margin)));

    for (int y = bounded.top(); y <= bounded.bottom(); ++y) {
        const double ty = bounded.height() <= 1
            ? 0.0
            : static_cast<double>(y - bounded.top()) / (bounded.height() - 1);
        for (int x = bounded.left(); x <= bounded.right(); ++x) {
            const double tx = bounded.width() <= 1
                ? 0.0
                : static_cast<double>(x - bounded.left()) / (bounded.width() - 1);

            QColor horizontal = mixColor(leftColor, rightColor, tx);
            QColor vertical = mixColor(topColor, bottomColor, ty);
            QColor mixed = mixColor(horizontal, vertical, 0.5);

            // If a strip is out of image bounds, its median becomes white. Blend back toward the
            // surrounding fallback color so edge selections do not create a harsh white patch.
            mixed = mixColor(mixed, fallback, 0.25);
            image.setPixelColor(x, y, mixed);
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
  <Application>Image2PPT</Application>
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
  <dc:title>Image2PPT Export</dc:title>
  <dc:creator>Image2PPT</dc:creator>
  <cp:lastModifiedBy>Image2PPT</cp:lastModifiedBy>
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
                <a:latin typeface="Microsoft YaHei"/>
                <a:ea typeface="Microsoft YaHei"/>
              </a:rPr>
              <a:t>%4</a:t>
            </a:r>
          </a:p>)").arg(fontSize).arg(bold).arg(colorToHex(region.textColor), xmlEscape(line));
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
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Image2PPT">
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
    <a:fontScheme name="Image2PPT">
      <a:majorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/><a:cs typeface="Arial"/></a:majorFont>
      <a:minorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/><a:cs typeface="Arial"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Image2PPT">
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
