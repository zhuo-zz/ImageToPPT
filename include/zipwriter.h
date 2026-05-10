#pragma once

#include <QByteArray>
#include <QFile>
#include <QString>
#include <QVector>

class ZipWriter
{
public:
    explicit ZipWriter(const QString &path);

    bool open(QString *errorMessage);
    bool addFile(const QString &name, const QByteArray &data, QString *errorMessage);
    bool close(QString *errorMessage);

private:
    struct Entry
    {
        QString name;
        quint32 crc = 0;
        quint32 size = 0;
        quint32 offset = 0;
    };

    static quint32 crc32(const QByteArray &data);
    void writeUInt16(quint16 value);
    void writeUInt32(quint32 value);

    QFile file_;
    QVector<Entry> entries_;
};
