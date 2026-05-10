#pragma once

#include "Region.h"

#include <QImage>
#include <QWidget>

class ImageCanvas : public QWidget {
    Q_OBJECT

public:
    explicit ImageCanvas(QWidget *parent = nullptr);

    bool loadImage(const QString &path);
    const QImage &image() const;
    const QVector<RegionItem> &regions() const;
    void setRegions(const QVector<RegionItem> &regions);
    void selectRegion(int index);
    void clearRegions();
    void removeSelectedRegion();
    int selectedIndex() const;

signals:
    void regionCreated(const QRect &imageRect);
    void selectionChanged(int index);
    void regionsChanged();

protected:
    void paintEvent(QPaintEvent *event) override;
    void mousePressEvent(QMouseEvent *event) override;
    void mouseMoveEvent(QMouseEvent *event) override;
    void mouseReleaseEvent(QMouseEvent *event) override;
    void resizeEvent(QResizeEvent *event) override;

private:
    QRect imageDisplayRect() const;
    QPoint widgetToImage(const QPoint &point) const;
    QRect widgetRectToImageRect(const QRect &rect) const;
    QRect imageRectToWidgetRect(const QRect &rect) const;
    int hitTestRegion(const QPoint &point) const;

private:
    QImage m_image;
    QVector<RegionItem> m_regions;
    bool m_dragging = false;
    QPoint m_dragStart;
    QPoint m_dragEnd;
    int m_selectedIndex = -1;
};
