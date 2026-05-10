#include "MainWindow.h"

#include "PptxExporter.h"
#include "RegionDialog.h"

#include <QAction>
#include <QDialog>
#include <QDialogButtonBox>
#include <QFileDialog>
#include <QHBoxLayout>
#include <QKeySequence>
#include <QLabel>
#include <QListWidget>
#include <QMessageBox>
#include <QPixmap>
#include <QPushButton>
#include <QScrollArea>
#include <QSignalBlocker>
#include <QStatusBar>
#include <QToolBar>
#include <QVBoxLayout>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
{
    setWindowTitle(QStringLiteral("ImageToPPT - 图片转 PPT 半自动重建工具"));

    m_canvas = new ImageCanvas(this);
    m_regionList = new QListWidget(this);
    m_regionList->setMinimumWidth(260);

    auto *editButton = new QPushButton(QStringLiteral("编辑选中区域"), this);
    auto *deleteButton = new QPushButton(QStringLiteral("删除选中区域"), this);

    auto *sideLayout = new QVBoxLayout;
    sideLayout->addWidget(m_regionList, 1);
    sideLayout->addWidget(editButton);
    sideLayout->addWidget(deleteButton);

    auto *side = new QWidget(this);
    side->setLayout(sideLayout);

    auto *mainLayout = new QHBoxLayout;
    mainLayout->addWidget(m_canvas, 1);
    mainLayout->addWidget(side);

    auto *central = new QWidget(this);
    central->setLayout(mainLayout);
    setCentralWidget(central);

    auto *openAction = new QAction(QStringLiteral("打开图片"), this);
    auto *previewAction = new QAction(QStringLiteral("导出预览"), this);
    auto *exportAction = new QAction(QStringLiteral("导出 PPTX"), this);
    auto *clearAction = new QAction(QStringLiteral("清空区域"), this);
    m_backgroundAction = new QAction(QStringLiteral("高保真背景"), this);
    m_backgroundAction->setCheckable(true);
    m_backgroundAction->setChecked(true);
    m_undoAction = new QAction(QStringLiteral("撤销"), this);
    m_undoAction->setShortcut(QKeySequence::Undo);
    m_undoAction->setEnabled(false);

    auto *toolbar = addToolBar(QStringLiteral("工具栏"));
    toolbar->addAction(openAction);
    toolbar->addAction(previewAction);
    toolbar->addAction(exportAction);
    toolbar->addAction(m_undoAction);
    toolbar->addAction(m_backgroundAction);
    toolbar->addAction(clearAction);

    connect(openAction, &QAction::triggered, this, &MainWindow::openImage);
    connect(previewAction, &QAction::triggered, this, &MainWindow::previewExport);
    connect(exportAction, &QAction::triggered, this, &MainWindow::exportPptx);
    connect(m_undoAction, &QAction::triggered, this, &MainWindow::undoLastChange);
    connect(clearAction, &QAction::triggered, this, &MainWindow::clearRegions);
    connect(m_canvas, &ImageCanvas::regionCreated, this, &MainWindow::addRegion);
    connect(m_canvas, &ImageCanvas::regionsChanged, this, &MainWindow::refreshRegionList);
    connect(m_canvas, &ImageCanvas::selectionChanged, this, [this](int index) {
        QSignalBlocker blocker(m_regionList);
        m_regionList->setCurrentRow(index);
    });
    connect(editButton, &QPushButton::clicked, this, &MainWindow::editSelectedRegion);
    connect(deleteButton, &QPushButton::clicked, this, &MainWindow::deleteSelectedRegion);
    connect(m_regionList, &QListWidget::itemSelectionChanged, this, &MainWindow::onListSelectionChanged);

    statusBar()->showMessage(QStringLiteral("打开图片后，用鼠标拖拽框选图片或文字区域"));
}

void MainWindow::openImage()
{
    const QString path = QFileDialog::getOpenFileName(
        this,
        QStringLiteral("选择图片"),
        QString(),
        QStringLiteral("Images (*.png *.jpg *.jpeg *.bmp)")
    );

    if (path.isEmpty())
        return;

    if (!m_canvas->loadImage(path)) {
        QMessageBox::warning(this, QStringLiteral("打开失败"), QStringLiteral("无法读取该图片。"));
        return;
    }

    m_currentImagePath = path;
    clearUndoHistory();
    statusBar()->showMessage(QStringLiteral("已打开图片：%1").arg(path));
}

void MainWindow::previewExport()
{
    showExportPreview(false);
}

void MainWindow::exportPptx()
{
    if (m_canvas->image().isNull()) {
        QMessageBox::information(this, QStringLiteral("提示"), QStringLiteral("请先打开一张图片。"));
        return;
    }

    if (m_canvas->regions().isEmpty()) {
        QMessageBox::information(this, QStringLiteral("提示"), QStringLiteral("请至少框选一个区域。"));
        return;
    }

    if (!showExportPreview(true))
        return;

    const QString path = QFileDialog::getSaveFileName(
        this,
        QStringLiteral("导出 PPTX"),
        QStringLiteral("output.pptx"),
        QStringLiteral("PowerPoint (*.pptx)")
    );

    if (path.isEmpty())
        return;

    PptxExporter exporter;
    QString error;
    if (!exporter.exportPptx(path,
                             m_canvas->image(),
                             m_canvas->regions(),
                             m_backgroundAction->isChecked(),
                             &error)) {
        QMessageBox::critical(this, QStringLiteral("导出失败"), error);
        return;
    }

    QMessageBox::information(this, QStringLiteral("导出成功"), QStringLiteral("PPTX 已生成：\n%1").arg(path));
}

bool MainWindow::showExportPreview(bool allowContinue)
{
    if (m_canvas->image().isNull()) {
        QMessageBox::information(this, QStringLiteral("提示"), QStringLiteral("请先打开一张图片。"));
        return false;
    }

    if (m_canvas->regions().isEmpty()) {
        QMessageBox::information(this, QStringLiteral("提示"), QStringLiteral("请至少框选一个区域。"));
        return false;
    }

    if (!warnIfComplexImageRegions(allowContinue))
        return false;

    PptxExporter exporter;
    const QImage preview = exporter.buildPreviewImage(
        m_canvas->image(),
        m_canvas->regions(),
        m_backgroundAction->isChecked()
    );
    if (preview.isNull()) {
        QMessageBox::warning(this, QStringLiteral("预览失败"), QStringLiteral("无法生成导出预览。"));
        return false;
    }

    QDialog dialog(this);
    dialog.setWindowTitle(QStringLiteral("导出前预览"));
    dialog.resize(980, 720);

    auto *imageLabel = new QLabel(&dialog);
    imageLabel->setAlignment(Qt::AlignCenter);
    imageLabel->setBackgroundRole(QPalette::Base);
    imageLabel->setSizePolicy(QSizePolicy::Ignored, QSizePolicy::Ignored);

    QPixmap pixmap = QPixmap::fromImage(preview);
    const QSize maxPreviewSize(900, 600);
    if (pixmap.width() > maxPreviewSize.width() || pixmap.height() > maxPreviewSize.height())
        pixmap = pixmap.scaled(maxPreviewSize, Qt::KeepAspectRatio, Qt::SmoothTransformation);
    imageLabel->setPixmap(pixmap);
    imageLabel->adjustSize();

    auto *scrollArea = new QScrollArea(&dialog);
    scrollArea->setWidget(imageLabel);
    scrollArea->setAlignment(Qt::AlignCenter);
    scrollArea->setWidgetResizable(false);

    auto *hintLabel = new QLabel(QStringLiteral("预览包含背景修复、图片区域和文字区域。请确认文字位置和背景效果后再导出。"), &dialog);
    hintLabel->setWordWrap(true);

    auto *buttons = new QDialogButtonBox(&dialog);
    if (allowContinue) {
        buttons->addButton(QStringLiteral("返回修改"), QDialogButtonBox::RejectRole);
        buttons->addButton(QStringLiteral("继续导出"), QDialogButtonBox::AcceptRole);
    } else {
        buttons->addButton(QStringLiteral("关闭"), QDialogButtonBox::AcceptRole);
    }

    auto *layout = new QVBoxLayout(&dialog);
    layout->addWidget(hintLabel);
    layout->addWidget(scrollArea, 1);
    layout->addWidget(buttons);

    connect(buttons, &QDialogButtonBox::accepted, &dialog, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, &dialog, &QDialog::reject);

    return dialog.exec() == QDialog::Accepted;
}

bool MainWindow::warnIfComplexImageRegions(bool allowContinue)
{
    QStringList complexRegions;
    const auto &regions = m_canvas->regions();
    for (int i = 0; i < regions.size(); ++i) {
        const RegionItem &region = regions.at(i);
        if (region.type != RegionType::Image)
            continue;

        if (!region.hasImageBackgroundColor
            && !PptxExporter::isSimpleBackgroundRegion(m_canvas->image(), region.imageRect))
            complexRegions.append(QStringLiteral("%1").arg(i + 1));
    }

    if (complexRegions.isEmpty())
        return true;

    const QString message = QStringLiteral(
        "检测到第 %1 个图片区域不是纯色或简单背景。\n\n"
        "当前版本只建议在纯色背景中处理图片区域；复杂背景不会自动补齐，否则容易产生明显涂抹痕迹。\n\n"
        "建议：把该区域改为普通保留背景，或只在纯色底图标、按钮、小块图片上使用图片区域。"
    ).arg(complexRegions.join(QStringLiteral("、")));

    QMessageBox::information(
        this,
        QStringLiteral("图片区域背景取色"),
        message + QStringLiteral("\n\n请先在背景中框选一小块颜色区域，程序会用该颜色填充图片区域底下的背景。")
    );
    requestImageBackgroundSample(complexRegions.first().toInt() - 1);
    return false;

    if (!allowContinue) {
        QMessageBox::information(this, QStringLiteral("图片区域提示"), message);
        return true;
    }

    const auto result = QMessageBox::warning(
        this,
        QStringLiteral("图片区域提示"),
        message + QStringLiteral("\n\n是否仍然继续导出？"),
        QMessageBox::Yes | QMessageBox::No,
        QMessageBox::No
    );
    return result == QMessageBox::Yes;
}

bool MainWindow::requestImageBackgroundSample(int regionIndex)
{
    const auto result = QMessageBox::information(
        this,
        QStringLiteral("图片区域背景取色"),
        QStringLiteral(
            "该图片区域不是纯色或简单背景。\n\n"
            "请在原图背景中框选一小块合适的背景颜色，程序会用该颜色填充图片区域底下的背景。\n\n"
            "建议选择靠近该图片区域、颜色较干净的小块背景。"
        ),
        QMessageBox::Ok | QMessageBox::Cancel,
        QMessageBox::Ok
    );

    if (result != QMessageBox::Ok)
        return false;

    m_pendingImageBackgroundSampleIndex = regionIndex;
    m_canvas->selectRegion(regionIndex);
    statusBar()->showMessage(QStringLiteral("请在画布上框选一小块背景颜色区域，用于填充图片区域背景。"));
    return true;
}

void MainWindow::applyImageBackgroundSample(const QRect &sampleRect)
{
    QVector<RegionItem> regions = m_canvas->regions();
    if (m_pendingImageBackgroundSampleIndex < 0
        || m_pendingImageBackgroundSampleIndex >= regions.size()) {
        m_pendingImageBackgroundSampleIndex = -1;
        return;
    }

    RegionItem &region = regions[m_pendingImageBackgroundSampleIndex];
    region.imageBackgroundColor = PptxExporter::sampleRegionColor(m_canvas->image(), sampleRect);
    region.hasImageBackgroundColor = true;

    const int updatedIndex = m_pendingImageBackgroundSampleIndex;
    m_pendingImageBackgroundSampleIndex = -1;

    pushUndoState();
    m_canvas->setRegions(regions);
    m_canvas->selectRegion(updatedIndex);
    refreshRegionList();
    m_regionList->setCurrentRow(updatedIndex);
    statusBar()->showMessage(QStringLiteral("已记录图片区域的背景取色，导出时会用该颜色补齐背景。"));
}

void MainWindow::addRegion(const QRect &imageRect)
{
    if (m_pendingImageBackgroundSampleIndex >= 0) {
        applyImageBackgroundSample(imageRect);
        return;
    }

    RegionItem item;
    item.imageRect = imageRect;

    RegionDialog dialog(this);
    dialog.setRegion(item);
    dialog.setOcrSource(m_canvas->image());
    if (dialog.exec() != QDialog::Accepted)
        return;

    QVector<RegionItem> regions = m_canvas->regions();
    const RegionItem region = dialog.region();
    if (region.type == RegionType::Image
        && false
        && !PptxExporter::isSimpleBackgroundRegion(m_canvas->image(), region.imageRect)) {
        QMessageBox::information(
            this,
            QStringLiteral("图片区域提示"),
            QStringLiteral("该图片区域背景较复杂。当前版本只在纯色或简单背景中补齐图片区域；复杂背景导出时不会自动补齐，以免产生明显涂抹痕迹。")
        );
    }
    regions.append(region);
    pushUndoState();
    m_canvas->setRegions(regions);

    const int newIndex = regions.size() - 1;
    if (region.type == RegionType::Image
        && !region.hasImageBackgroundColor
        && !PptxExporter::isSimpleBackgroundRegion(m_canvas->image(), region.imageRect))
        requestImageBackgroundSample(newIndex);
}

void MainWindow::editSelectedRegion()
{
    const int row = m_regionList->currentRow();
    if (row < 0 || row >= m_canvas->regions().size())
        return;

    QVector<RegionItem> regions = m_canvas->regions();
    RegionDialog dialog(this);
    dialog.setRegion(regions.at(row));
    dialog.setOcrSource(m_canvas->image());
    if (dialog.exec() != QDialog::Accepted)
        return;

    const RegionItem region = dialog.region();
    if (region.type == RegionType::Image
        && false
        && !PptxExporter::isSimpleBackgroundRegion(m_canvas->image(), region.imageRect)) {
        QMessageBox::information(
            this,
            QStringLiteral("图片区域提示"),
            QStringLiteral("该图片区域背景较复杂。当前版本只在纯色或简单背景中补齐图片区域；复杂背景导出时不会自动补齐，以免产生明显涂抹痕迹。")
        );
    }
    regions[row] = region;
    pushUndoState();
    m_canvas->setRegions(regions);
    m_regionList->setCurrentRow(row);

    if (region.type == RegionType::Image
        && !region.hasImageBackgroundColor
        && !PptxExporter::isSimpleBackgroundRegion(m_canvas->image(), region.imageRect))
        requestImageBackgroundSample(row);
}

void MainWindow::deleteSelectedRegion()
{
    if (m_canvas->selectedIndex() < 0)
        return;

    pushUndoState();
    m_canvas->removeSelectedRegion();
    updateUndoAction();
}

void MainWindow::clearRegions()
{
    if (m_canvas->regions().isEmpty())
        return;

    pushUndoState();
    m_canvas->clearRegions();
    updateUndoAction();
}

void MainWindow::undoLastChange()
{
    if (m_undoStack.isEmpty())
        return;

    const QVector<RegionItem> previous = m_undoStack.takeLast();
    m_canvas->setRegions(previous);
    updateUndoAction();
}

void MainWindow::refreshRegionList()
{
    const int oldRow = m_regionList->currentRow();
    m_regionList->clear();

    const auto &regions = m_canvas->regions();
    for (int i = 0; i < regions.size(); ++i)
        m_regionList->addItem(regionLabel(regions.at(i), i));

    if (oldRow >= 0 && oldRow < regions.size())
        m_regionList->setCurrentRow(oldRow);
}

void MainWindow::onListSelectionChanged()
{
    const int row = m_regionList->currentRow();
    m_canvas->selectRegion(row);
}

void MainWindow::pushUndoState()
{
    m_undoStack.append(m_canvas->regions());
    constexpr int maxUndoSteps = 50;
    while (m_undoStack.size() > maxUndoSteps)
        m_undoStack.removeFirst();
    updateUndoAction();
}

void MainWindow::clearUndoHistory()
{
    m_undoStack.clear();
    updateUndoAction();
}

void MainWindow::updateUndoAction()
{
    if (m_undoAction)
        m_undoAction->setEnabled(!m_undoStack.isEmpty());
}

QString MainWindow::regionLabel(const RegionItem &item, int index) const
{
    const QString type = item.type == RegionType::Image ? QStringLiteral("图片") : QStringLiteral("文字");
    return QStringLiteral("%1. %2 [%3,%4 %5x%6]")
        .arg(index + 1)
        .arg(type)
        .arg(item.imageRect.x())
        .arg(item.imageRect.y())
        .arg(item.imageRect.width())
        .arg(item.imageRect.height());
}
