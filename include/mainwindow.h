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
    void previewExport();
    void exportPptx();
    void addRegion(const QRect &imageRect);
    void editSelectedRegion();
    void deleteSelectedRegion();
    void clearRegions();
    void undoLastChange();
    void refreshRegionList();
    void onListSelectionChanged();

private:
    QString regionLabel(const RegionItem &item, int index) const;
    bool showExportPreview(bool allowContinue);
    bool warnIfComplexImageRegions(bool allowContinue);
    bool requestImageBackgroundSample(int regionIndex);
    void applyImageBackgroundSample(const QRect &sampleRect);
    void pushUndoState();
    void clearUndoHistory();
    void updateUndoAction();

private:
    ImageCanvas *m_canvas = nullptr;
    QListWidget *m_regionList = nullptr;
    QAction *m_undoAction = nullptr;
    QAction *m_backgroundAction = nullptr;
    QString m_currentImagePath;
    QVector<QVector<RegionItem>> m_undoStack;
    int m_pendingImageBackgroundSampleIndex = -1;
};
