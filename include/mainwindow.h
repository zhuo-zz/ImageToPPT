#pragma once

#include <QImage>
#include <QMainWindow>

class QLabel;
class QPushButton;
class QSpinBox;
class QCheckBox;

class PreviewWidget;

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = nullptr);

private slots:
    void chooseImage();
    void exportPptx();
    void updatePreview();

private:
    void updateControls();

    QImage image_;
    QString imagePath_;
    PreviewWidget *preview_ = nullptr;
    QLabel *infoLabel_ = nullptr;
    QSpinBox *rowsSpin_ = nullptr;
    QSpinBox *colsSpin_ = nullptr;
    QCheckBox *gapCheck_ = nullptr;
    QPushButton *exportButton_ = nullptr;
};
