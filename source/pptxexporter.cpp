#include "pptxexporter.h"

#include "zipwriter.h"

#include <QBuffer>
#include <QObject>
#include <QRect>
#include <QtMath>

namespace {

constexpr qint64 SlideWidth = 13'333'333;
constexpr qint64 SlideHeight = 7'500'000;

QByteArray xml(const QString &text)
{
    return text.toUtf8();
}

QString contentTypesXml()
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
}

QString rootRelsXml()
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
)");
}

QString presentationXml()
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId2"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
  <p:sldSz cx="%1" cy="%2" type="wide"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>
)").arg(SlideWidth).arg(SlideHeight);
}

QString presentationRelsXml()
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
</Relationships>
)");
}

QString themeXml()
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Image Split Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F2937"/></a:dk2>
      <a:lt2><a:srgbClr val="F8FAFC"/></a:lt2>
      <a:accent1><a:srgbClr val="2563EB"/></a:accent1>
      <a:accent2><a:srgbClr val="16A34A"/></a:accent2>
      <a:accent3><a:srgbClr val="F59E0B"/></a:accent3>
      <a:accent4><a:srgbClr val="DC2626"/></a:accent4>
      <a:accent5><a:srgbClr val="7C3AED"/></a:accent5>
      <a:accent6><a:srgbClr val="0891B2"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Arial"/></a:majorFont>
      <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>
)");
}

QString slideMasterXml()
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>
)");
}

QString slideMasterRelsXml()
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>
)");
}

QString slideLayoutXml()
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>
)");
}

QString slideLayoutRelsXml()
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>
)");
}

QString appXml(int tileCount)
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>ImageToEditablePptx</Application>
  <PresentationFormat>Widescreen</PresentationFormat>
  <Slides>1</Slides>
  <Notes>0</Notes>
  <HiddenSlides>0</HiddenSlides>
  <MMClips>0</MMClips>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Slides</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Split image, %1 units</vt:lpstr></vt:vector></TitlesOfParts>
</Properties>
)").arg(tileCount);
}

QString coreXml()
{
    return QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Split Image</dc:title>
  <dc:creator>ImageToEditablePptx</dc:creator>
  <cp:lastModifiedBy>ImageToEditablePptx</cp:lastModifiedBy>
</cp:coreProperties>
)");
}

QString slideRelsXml(int tileCount)
{
    QString rels = QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
)");
    for (int i = 1; i <= tileCount; ++i) {
        rels += QStringLiteral("  <Relationship Id=\"rId%1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/tile_%1.png\"/>\n").arg(i);
    }
    rels += QStringLiteral("</Relationships>\n");
    return rels;
}

QString slideXml(const QVector<QRect> &sourceRects, const QSize &imageSize, bool addSmallGaps)
{
    const double imageRatio = static_cast<double>(imageSize.width()) / imageSize.height();
    const double slideRatio = static_cast<double>(SlideWidth) / SlideHeight;

    qint64 fittedWidth = SlideWidth;
    qint64 fittedHeight = SlideHeight;
    if (imageRatio > slideRatio) {
        fittedHeight = qRound64(SlideWidth / imageRatio);
    } else {
        fittedWidth = qRound64(SlideHeight * imageRatio);
    }

    const qint64 originX = (SlideWidth - fittedWidth) / 2;
    const qint64 originY = (SlideHeight - fittedHeight) / 2;
    const qint64 gap = addSmallGaps ? 9144 : 0;

    QString body = QStringLiteral(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
)");

    for (int i = 0; i < sourceRects.size(); ++i) {
        const QRect rect = sourceRects.at(i);
        qint64 x = originX + qRound64(static_cast<double>(rect.x()) / imageSize.width() * fittedWidth);
        qint64 y = originY + qRound64(static_cast<double>(rect.y()) / imageSize.height() * fittedHeight);
        qint64 cx = qRound64(static_cast<double>(rect.width()) / imageSize.width() * fittedWidth);
        qint64 cy = qRound64(static_cast<double>(rect.height()) / imageSize.height() * fittedHeight);
        if (gap > 0 && cx > gap * 2 && cy > gap * 2) {
            x += gap / 2;
            y += gap / 2;
            cx -= gap;
            cy -= gap;
        }

        body += QStringLiteral(R"(      <p:pic>
        <p:nvPicPr><p:cNvPr id="%1" name="Tile %2"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rId%2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr><a:xfrm><a:off x="%3" y="%4"/><a:ext cx="%5" cy="%6"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      </p:pic>
)").arg(i + 2).arg(i + 1).arg(x).arg(y).arg(cx).arg(cy);
    }

    body += QStringLiteral(R"(    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
)");
    return body;
}

bool addXml(ZipWriter &zip, const QString &name, const QString &content, QString *errorMessage)
{
    return zip.addFile(name, xml(content), errorMessage);
}

} // namespace

bool PptxExporter::exportImageGrid(const QImage &image,
                                   const QString &outputPath,
                                   const PptxExportOptions &options,
                                   QString *errorMessage)
{
    if (image.isNull()) {
        if (errorMessage) {
            *errorMessage = QObject::tr("Image is empty.");
        }
        return false;
    }
    if (options.rows < 1 || options.columns < 1) {
        if (errorMessage) {
            *errorMessage = QObject::tr("Rows and columns must be greater than zero.");
        }
        return false;
    }

    QVector<QRect> rects;
    rects.reserve(options.rows * options.columns);
    for (int row = 0; row < options.rows; ++row) {
        const int y1 = image.height() * row / options.rows;
        const int y2 = image.height() * (row + 1) / options.rows;
        for (int column = 0; column < options.columns; ++column) {
            const int x1 = image.width() * column / options.columns;
            const int x2 = image.width() * (column + 1) / options.columns;
            rects.append(QRect(QPoint(x1, y1), QPoint(x2 - 1, y2 - 1)));
        }
    }

    ZipWriter zip(outputPath);
    if (!zip.open(errorMessage)) {
        return false;
    }

    const int tileCount = rects.size();
    if (!addXml(zip, QStringLiteral("[Content_Types].xml"), contentTypesXml(), errorMessage)
        || !addXml(zip, QStringLiteral("_rels/.rels"), rootRelsXml(), errorMessage)
        || !addXml(zip, QStringLiteral("docProps/app.xml"), appXml(tileCount), errorMessage)
        || !addXml(zip, QStringLiteral("docProps/core.xml"), coreXml(), errorMessage)
        || !addXml(zip, QStringLiteral("ppt/presentation.xml"), presentationXml(), errorMessage)
        || !addXml(zip, QStringLiteral("ppt/_rels/presentation.xml.rels"), presentationRelsXml(), errorMessage)
        || !addXml(zip, QStringLiteral("ppt/theme/theme1.xml"), themeXml(), errorMessage)
        || !addXml(zip, QStringLiteral("ppt/slideMasters/slideMaster1.xml"), slideMasterXml(), errorMessage)
        || !addXml(zip, QStringLiteral("ppt/slideMasters/_rels/slideMaster1.xml.rels"), slideMasterRelsXml(), errorMessage)
        || !addXml(zip, QStringLiteral("ppt/slideLayouts/slideLayout1.xml"), slideLayoutXml(), errorMessage)
        || !addXml(zip, QStringLiteral("ppt/slideLayouts/_rels/slideLayout1.xml.rels"), slideLayoutRelsXml(), errorMessage)
        || !addXml(zip, QStringLiteral("ppt/slides/slide1.xml"), slideXml(rects, image.size(), options.addSmallGaps), errorMessage)
        || !addXml(zip, QStringLiteral("ppt/slides/_rels/slide1.xml.rels"), slideRelsXml(tileCount), errorMessage)) {
        return false;
    }

    for (int i = 0; i < rects.size(); ++i) {
        const QImage tile = image.copy(rects.at(i));
        QByteArray bytes;
        QBuffer buffer(&bytes);
        buffer.open(QIODevice::WriteOnly);
        if (!tile.save(&buffer, "PNG")) {
            if (errorMessage) {
                *errorMessage = QObject::tr("Failed to encode an image tile.");
            }
            return false;
        }
        if (!zip.addFile(QStringLiteral("ppt/media/tile_%1.png").arg(i + 1), bytes, errorMessage)) {
            return false;
        }
    }

    return zip.close(errorMessage);
}
