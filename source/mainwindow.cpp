#include "MainWindow.h"

#include "PptxExporter.h"
#include "RegionDialog.h"

#include <QAction>
#include <QFileDialog>
#include <QHBoxLayout>
#include <QListWidget>
#include <QMessageBox>
#include <QPushButton>
#include <QSignalBlocker>
#include <QStatusBar>
#include <QToolBar>
#include <QVBoxLayout>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
{
    setWindowTitle(QStringLiteral("Image2PPT - 图片转 PPT 半自动重建工具"));

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
    auto *exportAction = new QAction(QStringLiteral("导出 PPTX"), this);
    auto *clearAction = new QAction(QStringLiteral("清空区域"), this);
    m_backgroundAction = new QAction(QStringLiteral("高保真背景"), this);
    m_backgroundAction->setCheckable(true);
    m_backgroundAction->setChecked(true);

    auto *toolbar = addToolBar(QStringLiteral("工具栏"));
    toolbar->addAction(openAction);
    toolbar->addAction(exportAction);
    toolbar->addAction(m_backgroundAction);
    toolbar->addAction(clearAction);

    connect(openAction, &QAction::triggered, this, &MainWindow::openImage);
    connect(exportAction, &QAction::triggered, this, &MainWindow::exportPptx);
    connect(clearAction, &QAction::triggered, m_canvas, &ImageCanvas::clearRegions);
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
    statusBar()->showMessage(QStringLiteral("已打开图片：%1").arg(path));
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

void MainWindow::addRegion(const QRect &imageRect)
{
    RegionItem item;
    item.imageRect = imageRect;

    RegionDialog dialog(this);
    dialog.setRegion(item);
    dialog.setOcrSource(m_canvas->image());
    if (dialog.exec() != QDialog::Accepted)
        return;

    QVector<RegionItem> regions = m_canvas->regions();
    regions.append(dialog.region());
    m_canvas->setRegions(regions);
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

    regions[row] = dialog.region();
    m_canvas->setRegions(regions);
    m_regionList->setCurrentRow(row);
}

void MainWindow::deleteSelectedRegion()
{
    m_canvas->removeSelectedRegion();
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
