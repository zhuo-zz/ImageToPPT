#include "mainwindow.h"

#include "pptxexporter.h"

#include <QCheckBox>
#include <QDialog>
#include <QFileDialog>
#include <QFileInfo>
#include <QHBoxLayout>
#include <QLabel>
#include <QMessageBox>
#include <QMouseEvent>
#include <QPainter>
#include <QPushButton>
#include <QSpinBox>
#include <QStatusBar>
#include <QTimer>
#include <QVBoxLayout>

class PreviewWidget : public QWidget
{
public:
    explicit PreviewWidget(QWidget *parent = nullptr) : QWidget(parent)
    {
        setMinimumSize(560, 420);
        setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Expanding);
        setCursor(Qt::PointingHandCursor);
    }

    void setImage(const QImage &image)
    {
        image_ = image;
        setCursor(image_.isNull() ? Qt::PointingHandCursor : Qt::ArrowCursor);
        repaint();
        update();
    }

    void setGrid(int rows, int columns)
    {
        rows_ = rows;
        columns_ = columns;
        update();
    }

protected:
    void paintEvent(QPaintEvent *) override
    {
        QPainter painter(this);
        painter.fillRect(rect(), QColor(248, 249, 251));
        painter.setRenderHint(QPainter::Antialiasing, true);

        if (image_.isNull()) {
            painter.setPen(QColor(92, 99, 112));
            painter.drawText(rect(), Qt::AlignCenter, tr("Click here or use Choose Image"));
            return;
        }

        const QSize availableSize = size() - QSize(40, 40);
        const QSize targetSize = image_.size().scaled(availableSize, Qt::KeepAspectRatio);
        const QRect imageRect((width() - targetSize.width()) / 2,
                              (height() - targetSize.height()) / 2,
                              targetSize.width(),
                              targetSize.height());

        painter.drawImage(imageRect, image_);
        painter.setPen(QPen(QColor(20, 27, 38), 2));
        painter.drawRect(imageRect.adjusted(0, 0, -1, -1));

        painter.setPen(QPen(QColor(0, 122, 204, 210), 1));
        for (int row = 1; row < rows_; ++row) {
            const int y = imageRect.top() + imageRect.height() * row / rows_;
            painter.drawLine(imageRect.left(), y, imageRect.right(), y);
        }
        for (int column = 1; column < columns_; ++column) {
            const int x = imageRect.left() + imageRect.width() * column / columns_;
            painter.drawLine(x, imageRect.top(), x, imageRect.bottom());
        }
    }

    void mouseReleaseEvent(QMouseEvent *event) override
    {
        if (image_.isNull() && event->button() == Qt::LeftButton) {
            QMetaObject::invokeMethod(window(), "chooseImage");
        }
        QWidget::mouseReleaseEvent(event);
    }

private:
    QImage image_;
    int rows_ = 3;
    int columns_ = 3;
};

MainWindow::MainWindow(QWidget *parent) : QMainWindow(parent)
{
    auto *central = new QWidget(this);
    auto *root = new QHBoxLayout(central);

    preview_ = new PreviewWidget(central);
    root->addWidget(preview_, 1);

    auto *panel = new QWidget(central);
    panel->setFixedWidth(310);
    auto *panelLayout = new QVBoxLayout(panel);
    panelLayout->setContentsMargins(18, 18, 18, 18);
    panelLayout->setSpacing(14);

    auto *title = new QLabel(tr("Image to PPTX units"), panel);
    QFont titleFont = title->font();
    titleFont.setPointSize(15);
    titleFont.setBold(true);
    title->setFont(titleFont);
    panelLayout->addWidget(title);

    auto *chooseButton = new QPushButton(tr("Choose Image"), panel);
    panelLayout->addWidget(chooseButton);

    infoLabel_ = new QLabel(tr("No image selected"), panel);
    infoLabel_->setWordWrap(true);
    panelLayout->addWidget(infoLabel_);

    rowsSpin_ = new QSpinBox(panel);
    rowsSpin_->setRange(1, 30);
    rowsSpin_->setValue(3);
    rowsSpin_->setPrefix(tr("Rows  "));
    panelLayout->addWidget(rowsSpin_);

    colsSpin_ = new QSpinBox(panel);
    colsSpin_->setRange(1, 30);
    colsSpin_->setValue(3);
    colsSpin_->setPrefix(tr("Columns  "));
    panelLayout->addWidget(colsSpin_);

    gapCheck_ = new QCheckBox(tr("Add small gaps on export"), panel);
    gapCheck_->setChecked(false);
    panelLayout->addWidget(gapCheck_);

    exportButton_ = new QPushButton(tr("Export PPTX"), panel);
    exportButton_->setEnabled(false);
    panelLayout->addWidget(exportButton_);
    panelLayout->addStretch(1);
    root->addWidget(panel);

    setCentralWidget(central);
    setWindowTitle(tr("Image To Editable PPTX"));
    statusBar()->showMessage(tr("Ready"));

    connect(chooseButton, &QPushButton::clicked, this, &MainWindow::chooseImage);
    connect(exportButton_, &QPushButton::clicked, this, &MainWindow::exportPptx);
    connect(rowsSpin_, qOverload<int>(&QSpinBox::valueChanged), this, &MainWindow::updatePreview);
    connect(colsSpin_, qOverload<int>(&QSpinBox::valueChanged), this, &MainWindow::updatePreview);
}

void MainWindow::chooseImage()
{
    QFileDialog dialog(this, tr("Choose Image"));
    dialog.setFileMode(QFileDialog::ExistingFile);
    dialog.setNameFilters({tr("Images (*.png *.jpg *.jpeg *.bmp *.webp)"), tr("All files (*.*)")});
    dialog.setOption(QFileDialog::DontUseNativeDialog, true);

    if (dialog.exec() != QDialog::Accepted) {
        return;
    }

    const QString path = dialog.selectedFiles().value(0);
    if (path.isEmpty()) {
        return;
    }

    QImage loaded(path);
    if (loaded.isNull()) {
        QMessageBox::warning(this, tr("Open failed"), tr("This image could not be loaded."));
        return;
    }

    image_ = loaded.convertToFormat(QImage::Format_ARGB32);
    imagePath_ = path;
    updateControls();
    updatePreview();
    QTimer::singleShot(0, preview_, qOverload<>(&QWidget::update));
    statusBar()->showMessage(tr("Loaded image: %1").arg(path), 6000);
}

void MainWindow::exportPptx()
{
    if (image_.isNull()) {
        return;
    }

    const QString output = QFileDialog::getSaveFileName(
        this,
        tr("Export PPTX"),
        QFileInfo(imagePath_).completeBaseName() + QStringLiteral("_split.pptx"),
        tr("PowerPoint files (*.pptx)"));
    if (output.isEmpty()) {
        return;
    }

    PptxExportOptions options;
    options.rows = rowsSpin_->value();
    options.columns = colsSpin_->value();
    options.addSmallGaps = gapCheck_->isChecked();

    QString error;
    if (!PptxExporter::exportImageGrid(image_, output, options, &error)) {
        QMessageBox::critical(this, tr("Export failed"), error);
        return;
    }

    statusBar()->showMessage(tr("Exported: %1").arg(output), 6000);
    QMessageBox::information(this,
                             tr("Export complete"),
                             tr("The PPTX has been created. Each tile is an independent selectable object."));
}

void MainWindow::updatePreview()
{
    preview_->setGrid(rowsSpin_->value(), colsSpin_->value());
}

void MainWindow::updateControls()
{
    infoLabel_->setText(tr("Selected: %1\nSize: %2 x %3 px\nClick Export PPTX to create the file.")
                            .arg(imagePath_)
                            .arg(image_.width())
                            .arg(image_.height()));
    exportButton_->setEnabled(!image_.isNull());
    preview_->setImage(image_);
}
