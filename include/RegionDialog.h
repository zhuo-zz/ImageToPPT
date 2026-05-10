#pragma once

#include "Region.h"

#include <QColor>
#include <QDialog>
#include <QImage>

class QCheckBox;
class QColor;
class QComboBox;
class QFontComboBox;
class QLabel;
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
    void autoFitBorder();
    void updateControls();

private:
    static QRect fitRectToBorder(const QImage &image, const QRect &rect);
    static int colorDistance(const QColor &a, const QColor &b);
    static int horizontalEdgeScore(const QImage &image, int y, int left, int right);
    static int verticalEdgeScore(const QImage &image, int x, int top, int bottom);
    void estimateTextStyleFromImage();
    void updateRectStatus();
    QString recognizeTextByTesseract();
    static QImage prepareOcrImage(const QImage &image, int mode);
    static QString cleanOcrText(const QString &text);
    static int ocrTextScore(const QString &text);
    static int ocrCandidateScore(const QString &text);
    static bool hasTesseractLanguage(const QStringList &languages, const QString &language);
    void showOcrDownloadHint(const QString &title, const QString &message, const QString &url);
    void setOcrStatus(const QString &text, bool warning = false);

private:
    QComboBox *m_typeCombo = nullptr;
    QTextEdit *m_textEdit = nullptr;
    QLabel *m_rectStatusLabel = nullptr;
    QLabel *m_ocrStatusLabel = nullptr;
    QFontComboBox *m_fontCombo = nullptr;
    QSpinBox *m_fontSizeSpin = nullptr;
    QCheckBox *m_boldCheck = nullptr;
    QPushButton *m_colorButton = nullptr;
    QPushButton *m_fitBorderButton = nullptr;
    QColor m_color = Qt::black;
    QColor m_imageBackgroundColor = Qt::white;
    QRect m_rect;
    QImage m_ocrImage;
    bool m_ocrTried = false;
    bool m_styleEstimated = false;
    bool m_ocrDownloadHintShown = false;
    bool m_hasImageBackgroundColor = false;
};
