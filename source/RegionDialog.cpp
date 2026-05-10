#include "RegionDialog.h"

#include <algorithm>

#include <QCheckBox>
#include <QColorDialog>
#include <QComboBox>
#include <QDesktopServices>
#include <QDialogButtonBox>
#include <QFontComboBox>
#include <QFormLayout>
#include <QLabel>
#include <QMessageBox>
#include <QPainter>
#include <QProcess>
#include <QPushButton>
#include <QSpinBox>
#include <QTemporaryDir>
#include <QTextEdit>
#include <QUrl>
#include <QVBoxLayout>

RegionDialog::RegionDialog(QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(QStringLiteral("编辑区域"));

    m_typeCombo = new QComboBox(this);
    m_typeCombo->addItem(QStringLiteral("图片区域"), static_cast<int>(RegionType::Image));
    m_typeCombo->addItem(QStringLiteral("文字区域"), static_cast<int>(RegionType::Text));

    m_fitBorderButton = new QPushButton(QStringLiteral("自动贴合边框"), this);
    m_rectStatusLabel = new QLabel(this);
    m_rectStatusLabel->setWordWrap(true);

    m_textEdit = new QTextEdit(this);
    m_textEdit->setPlaceholderText(QStringLiteral("输入要写入 PPT 的文字"));
    m_textEdit->setMinimumHeight(110);

    m_ocrStatusLabel = new QLabel(QStringLiteral("OCR：选择文字区域后自动尝试识别。"), this);
    m_ocrStatusLabel->setWordWrap(true);

    m_fontCombo = new QFontComboBox(this);
    m_fontCombo->setCurrentFont(QFont(QStringLiteral("Microsoft YaHei")));

    m_fontSizeSpin = new QSpinBox(this);
    m_fontSizeSpin->setRange(8, 96);
    m_fontSizeSpin->setValue(24);
    m_fontSizeSpin->setSuffix(QStringLiteral(" pt"));

    m_boldCheck = new QCheckBox(QStringLiteral("加粗"), this);

    m_colorButton = new QPushButton(QStringLiteral("选择颜色"), this);

    auto *form = new QFormLayout;
    form->addRow(QStringLiteral("边框"), m_fitBorderButton);
    form->addRow(QStringLiteral("范围"), m_rectStatusLabel);
    form->addRow(QStringLiteral("区域类型"), m_typeCombo);
    form->addRow(QStringLiteral("文字内容"), m_textEdit);
    form->addRow(QStringLiteral("OCR 状态"), m_ocrStatusLabel);
    form->addRow(QStringLiteral("字体"), m_fontCombo);
    form->addRow(QStringLiteral("字号"), m_fontSizeSpin);
    form->addRow(QStringLiteral("字重"), m_boldCheck);
    form->addRow(QStringLiteral("颜色"), m_colorButton);

    auto *buttons = new QDialogButtonBox(QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);

    auto *layout = new QVBoxLayout(this);
    layout->addLayout(form);
    layout->addWidget(buttons);

    connect(m_typeCombo, &QComboBox::currentIndexChanged, this, &RegionDialog::updateControls);
    connect(m_colorButton, &QPushButton::clicked, this, &RegionDialog::chooseColor);
    connect(m_fitBorderButton, &QPushButton::clicked, this, &RegionDialog::autoFitBorder);
    connect(buttons, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, this, &QDialog::reject);

    updateControls();
}

void RegionDialog::setRegion(const RegionItem &region)
{
    m_rect = region.imageRect;
    m_typeCombo->setCurrentIndex(region.type == RegionType::Image ? 0 : 1);
    m_textEdit->setPlainText(region.text);
    m_fontCombo->setCurrentFont(QFont(region.fontFamily.isEmpty() ? QStringLiteral("Microsoft YaHei") : region.fontFamily));
    m_fontSizeSpin->setValue(region.fontSize);
    m_boldCheck->setChecked(region.bold);
    m_color = region.textColor;
    m_hasImageBackgroundColor = region.hasImageBackgroundColor;
    m_imageBackgroundColor = region.imageBackgroundColor;
    m_styleEstimated = !region.text.trimmed().isEmpty();
    updateRectStatus();
    updateControls();
}

void RegionDialog::setOcrSource(const QImage &image)
{
    m_ocrImage = image;
    m_ocrTried = false;
    m_ocrDownloadHintShown = false;
    m_styleEstimated = false;
}

RegionItem RegionDialog::region() const
{
    RegionItem item;
    item.type = m_typeCombo->currentIndex() == 0 ? RegionType::Image : RegionType::Text;
    item.imageRect = m_rect;
    item.text = m_textEdit->toPlainText();
    item.fontFamily = m_fontCombo->currentFont().family();
    item.fontSize = m_fontSizeSpin->value();
    item.bold = m_boldCheck->isChecked();
    item.textColor = m_color;
    item.hasImageBackgroundColor = m_hasImageBackgroundColor;
    item.imageBackgroundColor = m_imageBackgroundColor;
    return item;
}

void RegionDialog::chooseColor()
{
    const QColor color = QColorDialog::getColor(m_color, this, QStringLiteral("选择文字颜色"));
    if (color.isValid())
        m_color = color;
}

void RegionDialog::autoFitBorder()
{
    if (m_ocrImage.isNull()) {
        updateRectStatus();
        return;
    }

    const QRect fitted = fitRectToBorder(m_ocrImage, m_rect);
    if (fitted.isValid() && fitted != m_rect)
        m_rect = fitted;

    updateRectStatus();
}

int RegionDialog::colorDistance(const QColor &a, const QColor &b)
{
    return qAbs(a.red() - b.red())
        + qAbs(a.green() - b.green())
        + qAbs(a.blue() - b.blue());
}

int RegionDialog::horizontalEdgeScore(const QImage &image, int y, int left, int right)
{
    if (y < 0 || y + 1 >= image.height())
        return 0;

    left = qBound(0, left, image.width() - 1);
    right = qBound(0, right, image.width() - 1);
    if (left >= right)
        return 0;

    const int step = qMax(1, (right - left + 1) / 260);
    int sum = 0;
    int count = 0;
    for (int x = left; x <= right; x += step) {
        sum += colorDistance(image.pixelColor(x, y), image.pixelColor(x, y + 1));
        ++count;
    }

    return count == 0 ? 0 : sum / count;
}

int RegionDialog::verticalEdgeScore(const QImage &image, int x, int top, int bottom)
{
    if (x < 0 || x + 1 >= image.width())
        return 0;

    top = qBound(0, top, image.height() - 1);
    bottom = qBound(0, bottom, image.height() - 1);
    if (top >= bottom)
        return 0;

    const int step = qMax(1, (bottom - top + 1) / 220);
    int sum = 0;
    int count = 0;
    for (int y = top; y <= bottom; y += step) {
        sum += colorDistance(image.pixelColor(x, y), image.pixelColor(x + 1, y));
        ++count;
    }

    return count == 0 ? 0 : sum / count;
}

QRect RegionDialog::fitRectToBorder(const QImage &image, const QRect &rect)
{
    const QRect bounded = rect.normalized().intersected(image.rect());
    if (image.isNull() || bounded.width() < 8 || bounded.height() < 8)
        return bounded;

    const int marginX = qBound(8, bounded.width() / 4, 80);
    const int marginY = qBound(8, bounded.height() / 2, 60);
    const QRect search = bounded.adjusted(-marginX, -marginY, marginX, marginY).intersected(image.rect());

    auto bestHorizontal = [&](int from, int to) {
        int bestY = from;
        int bestScore = -1;
        for (int y = from; y <= to; ++y) {
            const int score = horizontalEdgeScore(image, y, search.left(), search.right());
            if (score > bestScore) {
                bestScore = score;
                bestY = y;
            }
        }
        return QPair<int, int>(bestY, bestScore);
    };

    auto bestVertical = [&](int from, int to) {
        int bestX = from;
        int bestScore = -1;
        for (int x = from; x <= to; ++x) {
            const int score = verticalEdgeScore(image, x, search.top(), search.bottom());
            if (score > bestScore) {
                bestScore = score;
                bestX = x;
            }
        }
        return QPair<int, int>(bestX, bestScore);
    };

    const int centerY = bounded.center().y();
    const int centerX = bounded.center().x();
    const auto top = bestHorizontal(search.top(), qMax(search.top(), centerY - 1));
    const auto bottom = bestHorizontal(qMin(search.bottom() - 1, centerY), search.bottom() - 1);
    const auto left = bestVertical(search.left(), qMax(search.left(), centerX - 1));
    const auto right = bestVertical(qMin(search.right() - 1, centerX), search.right() - 1);

    const int minHorizontalScore = 22;
    const int minVerticalScore = 18;
    const int fittedLeft = left.second >= minVerticalScore ? left.first + 1 : bounded.left();
    const int fittedRight = right.second >= minVerticalScore ? right.first : bounded.right();
    const int fittedTop = top.second >= minHorizontalScore ? top.first + 1 : bounded.top();
    const int fittedBottom = bottom.second >= minHorizontalScore ? bottom.first : bounded.bottom();

    QRect fitted(QPoint(fittedLeft, fittedTop), QPoint(fittedRight, fittedBottom));
    fitted = fitted.normalized().intersected(image.rect());
    if (fitted.width() < 8 || fitted.height() < 8)
        return bounded;

    QVector<QColor> backgroundSamples;
    const int edgeDepth = qBound(3, qMin(bounded.width(), bounded.height()) / 18, 10);
    const int edgeStep = qMax(1, qMin(bounded.width(), bounded.height()) / 160);
    for (int y = bounded.top(); y <= bounded.bottom(); y += edgeStep) {
        for (int dx = 0; dx < edgeDepth; ++dx) {
            backgroundSamples.append(image.pixelColor(bounded.left() + dx, y));
            backgroundSamples.append(image.pixelColor(bounded.right() - dx, y));
        }
    }
    for (int x = bounded.left(); x <= bounded.right(); x += edgeStep) {
        for (int dy = 0; dy < edgeDepth; ++dy) {
            backgroundSamples.append(image.pixelColor(x, bounded.top() + dy));
            backgroundSamples.append(image.pixelColor(x, bounded.bottom() - dy));
        }
    }

    if (!backgroundSamples.isEmpty()) {
        QVector<int> reds;
        QVector<int> greens;
        QVector<int> blues;
        reds.reserve(backgroundSamples.size());
        greens.reserve(backgroundSamples.size());
        blues.reserve(backgroundSamples.size());
        for (const QColor &color : backgroundSamples) {
            reds.append(color.red());
            greens.append(color.green());
            blues.append(color.blue());
        }

        auto median = [](QVector<int> values) {
            std::sort(values.begin(), values.end());
            return values.at(values.size() / 2);
        };

        const QColor background(median(reds), median(greens), median(blues));
        const QRect scanArea = bounded.intersected(image.rect());
        const int foregroundThreshold = 86;
        const double foregroundRatio = 0.08;

        auto rowHasForeground = [&](int y) {
            int foreground = 0;
            int count = 0;
            const int step = qMax(1, scanArea.width() / 280);
            for (int x = scanArea.left(); x <= scanArea.right(); x += step) {
                ++count;
                if (colorDistance(image.pixelColor(x, y), background) >= foregroundThreshold)
                    ++foreground;
            }
            return count > 0 && static_cast<double>(foreground) / count >= foregroundRatio;
        };

        auto columnHasForeground = [&](int x) {
            int foreground = 0;
            int count = 0;
            const int step = qMax(1, scanArea.height() / 180);
            for (int y = scanArea.top(); y <= scanArea.bottom(); y += step) {
                ++count;
                if (colorDistance(image.pixelColor(x, y), background) >= foregroundThreshold)
                    ++foreground;
            }
            return count > 0 && static_cast<double>(foreground) / count >= foregroundRatio;
        };

        int tightLeft = scanArea.left();
        int tightRight = scanArea.right();
        int tightTop = scanArea.top();
        int tightBottom = scanArea.bottom();

        while (tightLeft < tightRight && !columnHasForeground(tightLeft))
            ++tightLeft;
        while (tightRight > tightLeft && !columnHasForeground(tightRight))
            --tightRight;
        while (tightTop < tightBottom && !rowHasForeground(tightTop))
            ++tightTop;
        while (tightBottom > tightTop && !rowHasForeground(tightBottom))
            --tightBottom;

        QRect tight(QPoint(tightLeft, tightTop), QPoint(tightRight, tightBottom));
        if (tight.width() >= 8 && tight.height() >= 8 && tight.intersects(bounded)) {
            fitted = tight.intersected(image.rect());
        }
    }

    return fitted;
}

void RegionDialog::updateRectStatus()
{
    if (!m_rectStatusLabel)
        return;

    m_rectStatusLabel->setText(QStringLiteral("[%1,%2 %3x%4]")
        .arg(m_rect.x())
        .arg(m_rect.y())
        .arg(m_rect.width())
        .arg(m_rect.height()));
}

void RegionDialog::estimateTextStyleFromImage()
{
    const QRect bounded = m_rect.intersected(m_ocrImage.rect());
    if (m_ocrImage.isNull() || bounded.width() < 4 || bounded.height() < 4)
        return;

    QVector<QColor> borderColors;
    const int edgeDepth = qBound(1, qMin(bounded.width(), bounded.height()) / 14, 5);
    const int step = qMax(1, qMin(bounded.width(), bounded.height()) / 120);
    for (int y = bounded.top(); y <= bounded.bottom(); y += step) {
        for (int dx = 0; dx < edgeDepth; ++dx) {
            borderColors.append(m_ocrImage.pixelColor(bounded.left() + dx, y));
            borderColors.append(m_ocrImage.pixelColor(bounded.right() - dx, y));
        }
    }
    for (int x = bounded.left(); x <= bounded.right(); x += step) {
        for (int dy = 0; dy < edgeDepth; ++dy) {
            borderColors.append(m_ocrImage.pixelColor(x, bounded.top() + dy));
            borderColors.append(m_ocrImage.pixelColor(x, bounded.bottom() - dy));
        }
    }

    auto medianColor = [](QVector<QColor> colors) {
        if (colors.isEmpty())
            return QColor(Qt::white);

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
    };

    const QColor background = medianColor(borderColors);
    QVector<QColor> foregroundColors;
    QVector<int> rowCounts(bounded.height(), 0);
    int minY = bounded.bottom();
    int maxY = bounded.top();
    int foregroundPixels = 0;

    for (int y = bounded.top(); y <= bounded.bottom(); ++y) {
        for (int x = bounded.left(); x <= bounded.right(); ++x) {
            const QColor pixel = m_ocrImage.pixelColor(x, y);
            if (colorDistance(pixel, background) < 72)
                continue;

            foregroundColors.append(pixel);
            ++rowCounts[y - bounded.top()];
            ++foregroundPixels;
            minY = qMin(minY, y);
            maxY = qMax(maxY, y);
        }
    }

    if (foregroundColors.isEmpty())
        return;

    const QColor textColor = medianColor(foregroundColors);
    m_color = textColor;

    int activeRows = 0;
    const int minRowPixels = qMax(2, bounded.width() / 80);
    for (int count : rowCounts) {
        if (count >= minRowPixels)
            ++activeRows;
    }

    const int textHeight = qMax(activeRows, maxY - minY + 1);
    const int estimatedPointSize = qBound(8, qRound(textHeight * 0.72), 96);
    m_fontSizeSpin->setValue(estimatedPointSize);

    const double foregroundDensity = static_cast<double>(foregroundPixels) / qMax(1, bounded.width() * qMax(1, textHeight));
    m_boldCheck->setChecked(foregroundDensity > 0.18);

    const bool mostlyLatin = std::all_of(m_textEdit->toPlainText().cbegin(), m_textEdit->toPlainText().cend(), [](const QChar &ch) {
        return ch.unicode() < 0x0100 || ch.isSpace();
    });
    if (mostlyLatin && !m_textEdit->toPlainText().trimmed().isEmpty())
        m_fontCombo->setCurrentFont(QFont(QStringLiteral("Arial")));
    else
        m_fontCombo->setCurrentFont(QFont(QStringLiteral("Microsoft YaHei")));
}

void RegionDialog::updateControls()
{
    const bool textMode = m_typeCombo->currentIndex() == 1;
    m_textEdit->setEnabled(textMode);
    m_fontCombo->setEnabled(textMode);
    m_fontSizeSpin->setEnabled(textMode);
    m_boldCheck->setEnabled(textMode);
    m_colorButton->setEnabled(textMode);

    if (!textMode) {
        setOcrStatus(QStringLiteral("OCR：当前为图片区域，不进行文字识别。"));
        return;
    }

    if (!m_styleEstimated && !m_ocrImage.isNull()) {
        estimateTextStyleFromImage();
        m_styleEstimated = true;
    }

    if (textMode && !m_ocrTried && m_textEdit->toPlainText().trimmed().isEmpty() && !m_ocrImage.isNull()) {
        m_ocrTried = true;
        setOcrStatus(QStringLiteral("OCR：正在识别，请稍候..."));
        const QString text = recognizeTextByTesseract();
        if (!text.trimmed().isEmpty()) {
            m_textEdit->setPlainText(text.trimmed());
            setOcrStatus(QStringLiteral("OCR：已自动识别，请核对文字是否准确。"));
        } else {
            m_textEdit->setPlaceholderText(QStringLiteral("未检测到可用 OCR。可安装 Tesseract，或在这里手动输入文字。"));
            setOcrStatus(QStringLiteral("OCR：未得到可靠结果，请手动输入或调整框选范围。"), true);
        }
    } else if (m_textEdit->toPlainText().trimmed().isEmpty()) {
        setOcrStatus(QStringLiteral("OCR：选择文字区域后自动尝试识别。"));
    } else {
        setOcrStatus(QStringLiteral("OCR：当前已有文字，可直接修改。"));
    }
}

QString RegionDialog::recognizeTextByTesseract()
{
    const QRect bounded = m_rect.intersected(m_ocrImage.rect());
    if (bounded.isEmpty())
        return {};

    QTemporaryDir tempDir;
    if (!tempDir.isValid())
        return {};

    const QImage cropped = m_ocrImage.copy(bounded);
    QStringList imagePaths;
    for (int mode = 0; mode < 7; ++mode) {
        const QString imagePath = tempDir.path() + QStringLiteral("/ocr_%1.png").arg(mode);
        if (prepareOcrImage(cropped, mode).save(imagePath, "PNG"))
            imagePaths.append(imagePath);
    }
    if (imagePaths.isEmpty())
        return {};

    auto runOcr = [&](const QStringList &arguments, bool *started = nullptr) -> QString {
        QProcess process;
        process.start(QStringLiteral("tesseract"), arguments);
        if (!process.waitForStarted(3000)) {
            if (started)
                *started = false;
            return {};
        }
        if (started)
            *started = true;
        if (!process.waitForFinished(15000))
            return {};
        if (process.exitCode() != 0)
            return {};
        return QString::fromUtf8(process.readAllStandardOutput());
    };

    QStringList languages;
    bool tesseractStarted = false;
    const QString languageOutput = runOcr({QStringLiteral("--list-langs")}, &tesseractStarted);
    if (!tesseractStarted) {
        setOcrStatus(QStringLiteral("OCR：未检测到 Tesseract，当前只能手动输入。"), true);
        showOcrDownloadHint(
            QStringLiteral("未检测到 Tesseract OCR"),
            QStringLiteral("用途：Tesseract 用于从框选图片区域中自动识别文字。\n\n未检测到 tesseract 命令，当前仍可手动输入文字。安装后请确保 tesseract.exe 已加入 PATH。"),
            QStringLiteral("https://tesseract-ocr.github.io/tessdoc/Installation.html")
        );
        return {};
    }

    for (const QString &line : languageOutput.split('\n')) {
        const QString language = line.trimmed();
        if (!language.isEmpty() && !language.startsWith(QStringLiteral("List of")))
            languages.append(language);
    }

    const bool hasChinese = hasTesseractLanguage(languages, QStringLiteral("chi_sim"));
    if (!hasChinese) {
        setOcrStatus(QStringLiteral("OCR：未安装简体中文训练数据，中文识别可能不准确。"), true);
        showOcrDownloadHint(
            QStringLiteral("未安装中文 OCR 训练数据"),
            QStringLiteral("用途：chi_sim.traineddata 用于识别简体中文，例如标题、说明文字和中文段落。\n\n这是可选项：不安装也可以继续手动输入；安装后中文 OCR 准确率会明显提升。下载后请把 chi_sim.traineddata 放入 Tesseract 的 tessdata 目录。"),
            QStringLiteral("https://github.com/tesseract-ocr/tessdata/blob/main/chi_sim.traineddata")
        );
    }

    QStringList languageModes;
    if (hasChinese && hasTesseractLanguage(languages, QStringLiteral("eng"))) {
        languageModes << QStringLiteral("chi_sim+eng") << QStringLiteral("chi_sim");
    } else if (hasChinese) {
        languageModes << QStringLiteral("chi_sim");
    } else if (hasTesseractLanguage(languages, QStringLiteral("eng"))) {
        languageModes << QStringLiteral("eng");
    } else {
        languageModes << QStringLiteral("chi_sim+eng") << QStringLiteral("eng");
    }

    QString bestText;
    int bestScore = 0;
    const bool looksLikeSingleLine = cropped.width() > cropped.height() * 3;
    const QStringList pageSegModes = looksLikeSingleLine
        ? QStringList{QStringLiteral("7"), QStringLiteral("13"), QStringLiteral("6"), QStringLiteral("8")}
        : QStringList{QStringLiteral("6"), QStringLiteral("11"), QStringLiteral("4"), QStringLiteral("7")};

    for (const QString &languageMode : languageModes) {
        for (const QString &imagePath : imagePaths) {
            for (const QString &pageSegMode : pageSegModes) {
                const QString text = cleanOcrText(runOcr({
                    imagePath,
                    QStringLiteral("stdout"),
                    QStringLiteral("-l"),
                    languageMode,
                    QStringLiteral("--psm"),
                    pageSegMode,
                    QStringLiteral("--oem"),
                    QStringLiteral("1"),
                    QStringLiteral("-c"),
                    QStringLiteral("preserve_interword_spaces=1"),
                    QStringLiteral("-c"),
                    QStringLiteral("textord_heavy_nr=1")
                }));
                const int score = ocrCandidateScore(text);
                if (score > bestScore) {
                    bestScore = score;
                    bestText = text;
                }
                if (score >= 80)
                    return bestText;
            }
        }
    }

    const bool bestHasCjk = std::any_of(bestText.cbegin(), bestText.cend(), [](const QChar &ch) {
        const ushort unicode = ch.unicode();
        return (unicode >= 0x4E00 && unicode <= 0x9FFF) || (unicode >= 0x3400 && unicode <= 0x4DBF);
    });
    const int minimumScore = bestHasCjk ? 14 : 28;
    return bestScore >= minimumScore ? bestText : QString();
}

QImage RegionDialog::prepareOcrImage(const QImage &image, int mode)
{
    if (image.isNull())
        return {};

    const int scale = image.width() < 260 || image.height() < 80 ? 6 : 3;
    QImage scaled = image.scaled(image.size() * scale, Qt::KeepAspectRatio, Qt::SmoothTransformation)
        .convertToFormat(QImage::Format_RGB32);

    QImage output(scaled.size() + QSize(32, 32), QImage::Format_RGB32);
    output.fill(Qt::white);

    QPainter painter(&output);
    painter.drawImage(16, 16, scaled);
    painter.end();

    QVector<int> values;
    values.reserve(output.width() * output.height());
    for (int y = 0; y < output.height(); ++y) {
        for (int x = 0; x < output.width(); ++x) {
            const QColor color = output.pixelColor(x, y);
            values.append(qRound(color.red() * 0.299 + color.green() * 0.587 + color.blue() * 0.114));
        }
    }
    std::sort(values.begin(), values.end());
    const int low = values.at(values.size() / 12);
    const int high = values.at(values.size() * 11 / 12);
    const int range = qMax(1, high - low);

    for (int y = 0; y < output.height(); ++y) {
        for (int x = 0; x < output.width(); ++x) {
            const QColor color = output.pixelColor(x, y);
            int gray = qRound(color.red() * 0.299 + color.green() * 0.587 + color.blue() * 0.114);
            gray = qBound(0, (gray - low) * 255 / range, 255);
            if (mode == 0) {
                gray = qBound(0, qRound((gray - 128) * 1.55 + 128), 255);
            } else if (mode == 1) {
                gray = gray < 148 ? 0 : 255;
            } else if (mode == 2) {
                gray = gray < 180 ? 0 : 255;
            } else if (mode == 3) {
                gray = gray < 112 ? 0 : 255;
            } else if (mode == 4) {
                gray = gray > 96 ? 0 : 255;
            } else if (mode == 5) {
                gray = qBound(0, qRound((gray - 128) * 2.1 + 128), 255);
            } else {
                gray = gray < 205 ? 0 : 255;
            }
            output.setPixelColor(x, y, QColor(gray, gray, gray));
        }
    }

    return output;
}

QString RegionDialog::cleanOcrText(const QString &text)
{
    QString cleaned = text;
    cleaned.replace('\r', '\n');
    QStringList lines;
    for (QString line : cleaned.split('\n')) {
        line = line.simplified();
        if (!line.isEmpty())
            lines.append(line);
    }
    return lines.join('\n').trimmed();
}

int RegionDialog::ocrTextScore(const QString &text)
{
    if (text.trimmed().isEmpty())
        return 0;

    int score = 0;
    int cjkCount = 0;
    int alphaCount = 0;
    int digitCount = 0;
    int noisyCount = 0;

    for (const QChar ch : text) {
        const ushort unicode = ch.unicode();
        if ((unicode >= 0x4E00 && unicode <= 0x9FFF) || (unicode >= 0x3400 && unicode <= 0x4DBF)) {
            score += 8;
            ++cjkCount;
        } else if (ch.isLetter()) {
            score += 2;
            ++alphaCount;
        } else if (ch.isDigit()) {
            score += 2;
            ++digitCount;
        } else if (ch.isSpace()) {
            score += 1;
        } else if (QStringLiteral(".,:;!?()[]{}+-/·•、。，：；！？（）《》").contains(ch)) {
            score += 1;
        } else {
            score -= 4;
            ++noisyCount;
        }
    }

    if (cjkCount >= 2)
        score += 20;
    if (cjkCount == 0 && alphaCount <= 3 && digitCount == 0)
        score -= 18;
    if (noisyCount > qMax(2, text.size() / 4))
        score -= 20;

    return score;
}

int RegionDialog::ocrCandidateScore(const QString &text)
{
    if (text.trimmed().isEmpty())
        return 0;

    int score = ocrTextScore(text);
    int cjkCount = 0;
    int alphaCount = 0;
    int digitCount = 0;
    int noisyCount = 0;
    int latinWordCount = 0;
    int plausibleLatinWordCount = 0;
    QString currentLatinWord;

    auto finishLatinWord = [&]() {
        if (currentLatinWord.isEmpty())
            return;

        ++latinWordCount;
        bool hasVowel = false;
        for (const QChar ch : currentLatinWord) {
            if (QStringLiteral("AEIOUaeiou").contains(ch)) {
                hasVowel = true;
                break;
            }
        }

        if ((currentLatinWord.size() >= 4 && hasVowel) || currentLatinWord.size() >= 6) {
            ++plausibleLatinWordCount;
            score += 10;
        }
        currentLatinWord.clear();
    };

    for (const QChar ch : text) {
        const ushort unicode = ch.unicode();
        if ((unicode >= 0x4E00 && unicode <= 0x9FFF) || (unicode >= 0x3400 && unicode <= 0x4DBF)) {
            finishLatinWord();
            ++cjkCount;
        } else if (ch.isLetter()) {
            ++alphaCount;
            if (unicode < 0x0080)
                currentLatinWord.append(ch);
            else
                finishLatinWord();
        } else if (ch.isDigit()) {
            finishLatinWord();
            ++digitCount;
        } else if (ch.isSpace() || QStringLiteral(".,:;!?()[]{}+-/").contains(ch)) {
            finishLatinWord();
        } else {
            finishLatinWord();
            ++noisyCount;
        }
    }
    finishLatinWord();

    if (cjkCount >= 4)
        score += 22;
    if (latinWordCount > 0 && plausibleLatinWordCount == 0)
        score -= 18;
    if (cjkCount == 0 && alphaCount > 0 && alphaCount < 6)
        score -= 12;
    if (cjkCount == 0 && alphaCount > 0 && plausibleLatinWordCount == 0 && digitCount == 0)
        score -= 12;
    if (noisyCount > qMax(2, text.size() / 5))
        score -= 20;

    return score;
}

bool RegionDialog::hasTesseractLanguage(const QStringList &languages, const QString &language)
{
    return languages.contains(language, Qt::CaseInsensitive);
}

void RegionDialog::showOcrDownloadHint(const QString &title, const QString &message, const QString &url)
{
    if (m_ocrDownloadHintShown)
        return;

    m_ocrDownloadHintShown = true;

    QMessageBox box(this);
    box.setWindowTitle(title);
    box.setIcon(QMessageBox::Information);
    box.setText(message);
    box.setInformativeText(QStringLiteral("是否打开下载/安装说明页面？"));
    QPushButton *openButton = box.addButton(QStringLiteral("打开下载页"), QMessageBox::AcceptRole);
    box.addButton(QStringLiteral("稍后手动输入"), QMessageBox::RejectRole);
    box.exec();

    if (box.clickedButton() == openButton)
        QDesktopServices::openUrl(QUrl(url));
}

void RegionDialog::setOcrStatus(const QString &text, bool warning)
{
    if (!m_ocrStatusLabel)
        return;

    m_ocrStatusLabel->setText(text);
    m_ocrStatusLabel->setStyleSheet(warning
        ? QStringLiteral("color: #b45309;")
        : QStringLiteral("color: #4b5563;"));
}
