#pragma once

#include "ImageCanvas.h"

#include <QMainWindow>

class QAction;
class QListWidget;

class MainWindow : public QMainWindow {
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = nullptr);

private slots:
    void openImage();
    void exportPptx();
    void addRegion(const QRect &imageRect);
    void editSelectedRegion();
    void deleteSelectedRegion();
    void refreshRegionList();
    void onListSelectionChanged();

private:
    QString regionLabel(const RegionItem &item, int index) const;

private:
    ImageCanvas *m_canvas = nullptr;
    QListWidget *m_regionList = nullptr;
    QAction *m_backgroundAction = nullptr;
    QString m_currentImagePath;
};
