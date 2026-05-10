#include "ImageCanvas.h"

#include <QMouseEvent>
#include <QPainter>

ImageCanvas::ImageCanvas(QWidget *parent)
    : QWidget(parent)
{
    setMinimumSize(640, 420);
    setMouseTracking(true);
}

bool ImageCanvas::loadImage(const QString &path)
{
    QImage image(path);
    if (image.isNull())
        return false;

    m_image = image.convertToFormat(QImage::Format_ARGB32);
    clearRegions();
    update();
    return true;
}

const QImage &ImageCanvas::image() const
{
    return m_image;
}

const QVector<RegionItem> &ImageCanvas::regions() const
{
    return m_regions;
}

void ImageCanvas::setRegions(const QVector<RegionItem> &regions)
{
    m_regions = regions;
    m_selectedIndex = -1;
    update();
    emit regionsChanged();
    emit selectionChanged(m_selectedIndex);
}

void ImageCanvas::selectRegion(int index)
{
    if (index < -1 || index >= m_regions.size())
        return;

    m_selectedIndex = index;
    update();
    emit selectionChanged(m_selectedIndex);
}

void ImageCanvas::clearRegions()
{
    m_regions.clear();
    m_selectedIndex = -1;
    update();
    emit regionsChanged();
    emit selectionChanged(m_selectedIndex);
}

void ImageCanvas::removeSelectedRegion()
{
    if (m_selectedIndex < 0 || m_selectedIndex >= m_regions.size())
        return;

    m_regions.removeAt(m_selectedIndex);
    m_selectedIndex = -1;
    update();
    emit regionsChanged();
    emit selectionChanged(m_selectedIndex);
}

int ImageCanvas::selectedIndex() const
{
    return m_selectedIndex;
}

void ImageCanvas::paintEvent(QPaintEvent *)
{
    QPainter painter(this);
    painter.fillRect(rect(), QColor(245, 247, 250));

    if (m_image.isNull()) {
        painter.setPen(QColor(110, 118, 129));
        painter.drawText(rect(), Qt::AlignCenter, QStringLiteral("点击“打开图片”开始"));
        return;
    }

    const QRect displayRect = imageDisplayRect();
    painter.drawImage(displayRect, m_image);

    painter.setRenderHint(QPainter::Antialiasing);

    for (int i = 0; i < m_regions.size(); ++i) {
        const RegionItem &region = m_regions.at(i);
        const QRect wr = imageRectToWidgetRect(region.imageRect);
        const bool selected = (i == m_selectedIndex);
        const QColor base = region.type == RegionType::Image
            ? QColor(43, 127, 255)
            : QColor(36, 158, 95);

        QColor fill = base;
        fill.setAlpha(selected ? 70 : 34);
        painter.fillRect(wr, fill);

        QPen pen(base, selected ? 3 : 2);
        painter.setPen(pen);
        painter.drawRect(wr.adjusted(0, 0, -1, -1));

        painter.setPen(Qt::white);
        painter.fillRect(QRect(wr.left(), wr.top(), 54, 22), base);
        painter.drawText(QRect(wr.left(), wr.top(), 54, 22),
                         Qt::AlignCenter,
                         region.type == RegionType::Image ? QStringLiteral("图片") : QStringLiteral("文字"));
    }

    if (m_dragging) {
        QRect dragRect = QRect(m_dragStart, m_dragEnd).normalized();
        painter.fillRect(dragRect, QColor(255, 193, 7, 45));
        painter.setPen(QPen(QColor(255, 152, 0), 2, Qt::DashLine));
        painter.drawRect(dragRect.adjusted(0, 0, -1, -1));
    }
}

void ImageCanvas::mousePressEvent(QMouseEvent *event)
{
    if (m_image.isNull() || event->button() != Qt::LeftButton)
        return;

    const int hit = hitTestRegion(event->pos());
    if (hit >= 0) {
        m_selectedIndex = hit;
        emit selectionChanged(m_selectedIndex);
        update();
        return;
    }

    if (!imageDisplayRect().contains(event->pos()))
        return;

    m_dragging = true;
    m_dragStart = event->pos();
    m_dragEnd = event->pos();
    update();
}

void ImageCanvas::mouseMoveEvent(QMouseEvent *event)
{
    if (!m_dragging)
        return;

    m_dragEnd = event->pos();
    update();
}

void ImageCanvas::mouseReleaseEvent(QMouseEvent *event)
{
    if (!m_dragging || event->button() != Qt::LeftButton)
        return;

    m_dragging = false;
    m_dragEnd = event->pos();

    QRect imageRect = widgetRectToImageRect(QRect(m_dragStart, m_dragEnd).normalized());
    imageRect = imageRect.intersected(m_image.rect());

    if (imageRect.width() >= 8 && imageRect.height() >= 8)
        emit regionCreated(imageRect);

    update();
}

void ImageCanvas::resizeEvent(QResizeEvent *)
{
    update();
}

QRect ImageCanvas::imageDisplayRect() const
{
    if (m_image.isNull())
        return {};

    QSize scaled = m_image.size();
    scaled.scale(size() - QSize(24, 24), Qt::KeepAspectRatio);
    QPoint topLeft((width() - scaled.width()) / 2, (height() - scaled.height()) / 2);
    return QRect(topLeft, scaled);
}

QPoint ImageCanvas::widgetToImage(const QPoint &point) const
{
    const QRect displayRect = imageDisplayRect();
    if (displayRect.isEmpty())
        return {};

    const double sx = static_cast<double>(m_image.width()) / displayRect.width();
    const double sy = static_cast<double>(m_image.height()) / displayRect.height();

    return QPoint(
        qRound((point.x() - displayRect.x()) * sx),
        qRound((point.y() - displayRect.y()) * sy)
    );
}

QRect ImageCanvas::widgetRectToImageRect(const QRect &rect) const
{
    const QPoint p1 = widgetToImage(rect.topLeft());
    const QPoint p2 = widgetToImage(rect.bottomRight());
    return QRect(p1, p2).normalized();
}

QRect ImageCanvas::imageRectToWidgetRect(const QRect &rect) const
{
    const QRect displayRect = imageDisplayRect();
    if (displayRect.isEmpty())
        return {};

    const double sx = static_cast<double>(displayRect.width()) / m_image.width();
    const double sy = static_cast<double>(displayRect.height()) / m_image.height();

    return QRect(
        displayRect.x() + qRound(rect.x() * sx),
        displayRect.y() + qRound(rect.y() * sy),
        qRound(rect.width() * sx),
        qRound(rect.height() * sy)
    );
}

int ImageCanvas::hitTestRegion(const QPoint &point) const
{
    for (int i = m_regions.size() - 1; i >= 0; --i) {
        if (imageRectToWidgetRect(m_regions.at(i).imageRect).contains(point))
            return i;
    }
    return -1;
}
