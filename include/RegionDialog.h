#pragma once

#include "Region.h"

#include <QColor>
#include <QDialog>
#include <QImage>

class QCheckBox;
class QColor;
class QComboBox;
class QLineEdit;
class QPushButton;
class QSpinBox;
class QTextEdit;

class RegionDialog : public QDialog {
    Q_OBJECT

public:
    explicit RegionDialog(QWidget *parent = nullptr);

    void setRegion(const RegionItem &region);
    void setOcrSource(const QImage &image);
    RegionItem region() const;

private slots:
    void chooseColor();
    void updateControls();

private:
    QString recognizeTextByTesseract() const;

private:
    QComboBox *m_typeCombo = nullptr;
    QTextEdit *m_textEdit = nullptr;
    QSpinBox *m_fontSizeSpin = nullptr;
    QCheckBox *m_boldCheck = nullptr;
    QPushButton *m_colorButton = nullptr;
    QColor m_color = Qt::black;
    QRect m_rect;
    QImage m_ocrImage;
    bool m_ocrTried = false;
};
