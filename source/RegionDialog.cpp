#include "RegionDialog.h"

#include <QCheckBox>
#include <QColorDialog>
#include <QComboBox>
#include <QDialogButtonBox>
#include <QFormLayout>
#include <QProcess>
#include <QPushButton>
#include <QSpinBox>
#include <QTemporaryDir>
#include <QTextEdit>
#include <QVBoxLayout>

RegionDialog::RegionDialog(QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle(QStringLiteral("编辑区域"));

    m_typeCombo = new QComboBox(this);
    m_typeCombo->addItem(QStringLiteral("图片区域"), static_cast<int>(RegionType::Image));
    m_typeCombo->addItem(QStringLiteral("文字区域"), static_cast<int>(RegionType::Text));

    m_textEdit = new QTextEdit(this);
    m_textEdit->setPlaceholderText(QStringLiteral("输入要写入 PPT 的文字"));
    m_textEdit->setMinimumHeight(110);

    m_fontSizeSpin = new QSpinBox(this);
    m_fontSizeSpin->setRange(8, 96);
    m_fontSizeSpin->setValue(24);
    m_fontSizeSpin->setSuffix(QStringLiteral(" pt"));

    m_boldCheck = new QCheckBox(QStringLiteral("加粗"), this);

    m_colorButton = new QPushButton(QStringLiteral("选择颜色"), this);

    auto *form = new QFormLayout;
    form->addRow(QStringLiteral("区域类型"), m_typeCombo);
    form->addRow(QStringLiteral("文字内容"), m_textEdit);
    form->addRow(QStringLiteral("字号"), m_fontSizeSpin);
    form->addRow(QStringLiteral("字重"), m_boldCheck);
    form->addRow(QStringLiteral("颜色"), m_colorButton);

    auto *buttons = new QDialogButtonBox(QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);

    auto *layout = new QVBoxLayout(this);
    layout->addLayout(form);
    layout->addWidget(buttons);

    connect(m_typeCombo, &QComboBox::currentIndexChanged, this, &RegionDialog::updateControls);
    connect(m_colorButton, &QPushButton::clicked, this, &RegionDialog::chooseColor);
    connect(buttons, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, this, &QDialog::reject);

    updateControls();
}

void RegionDialog::setRegion(const RegionItem &region)
{
    m_rect = region.imageRect;
    m_typeCombo->setCurrentIndex(region.type == RegionType::Image ? 0 : 1);
    m_textEdit->setPlainText(region.text);
    m_fontSizeSpin->setValue(region.fontSize);
    m_boldCheck->setChecked(region.bold);
    m_color = region.textColor;
    updateControls();
}

void RegionDialog::setOcrSource(const QImage &image)
{
    m_ocrImage = image;
    m_ocrTried = false;
}

RegionItem RegionDialog::region() const
{
    RegionItem item;
    item.type = m_typeCombo->currentIndex() == 0 ? RegionType::Image : RegionType::Text;
    item.imageRect = m_rect;
    item.text = m_textEdit->toPlainText();
    item.fontSize = m_fontSizeSpin->value();
    item.bold = m_boldCheck->isChecked();
    item.textColor = m_color;
    return item;
}

void RegionDialog::chooseColor()
{
    const QColor color = QColorDialog::getColor(m_color, this, QStringLiteral("选择文字颜色"));
    if (color.isValid())
        m_color = color;
}

void RegionDialog::updateControls()
{
    const bool textMode = m_typeCombo->currentIndex() == 1;
    m_textEdit->setEnabled(textMode);
    m_fontSizeSpin->setEnabled(textMode);
    m_boldCheck->setEnabled(textMode);
    m_colorButton->setEnabled(textMode);

    if (textMode && !m_ocrTried && m_textEdit->toPlainText().trimmed().isEmpty() && !m_ocrImage.isNull()) {
        m_ocrTried = true;
        const QString text = recognizeTextByTesseract();
        if (!text.trimmed().isEmpty())
            m_textEdit->setPlainText(text.trimmed());
        else
            m_textEdit->setPlaceholderText(QStringLiteral("未检测到可用 OCR。可安装 Tesseract，或在这里手动输入文字。"));
    }
}

QString RegionDialog::recognizeTextByTesseract() const
{
    const QRect bounded = m_rect.intersected(m_ocrImage.rect());
    if (bounded.isEmpty())
        return {};

    QTemporaryDir tempDir;
    if (!tempDir.isValid())
        return {};

    const QString imagePath = tempDir.path() + QStringLiteral("/ocr_region.png");
    if (!m_ocrImage.copy(bounded).save(imagePath, "PNG"))
        return {};

    auto runOcr = [&](const QStringList &arguments) -> QString {
        QProcess process;
        process.start(QStringLiteral("tesseract"), arguments);
        if (!process.waitForStarted(3000))
            return {};
        if (!process.waitForFinished(15000))
            return {};
        if (process.exitCode() != 0)
            return {};
        return QString::fromUtf8(process.readAllStandardOutput());
    };

    QString text = runOcr({imagePath, QStringLiteral("stdout"), QStringLiteral("-l"),
                           QStringLiteral("chi_sim+eng"), QStringLiteral("--psm"), QStringLiteral("6")});
    if (!text.trimmed().isEmpty())
        return text;

    return runOcr({imagePath, QStringLiteral("stdout"), QStringLiteral("-l"),
                   QStringLiteral("eng"), QStringLiteral("--psm"), QStringLiteral("6")});
}
