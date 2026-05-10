#include "zipwriter.h"

#include <QObject>

ZipWriter::ZipWriter(const QString &path) : file_(path)
{
}

bool ZipWriter::open(QString *errorMessage)
{
    if (!file_.open(QIODevice::WriteOnly)) {
        if (errorMessage) {
            *errorMessage = QObject::tr("Failed to open output file: %1").arg(file_.errorString());
        }
        return false;
    }
    return true;
}

bool ZipWriter::addFile(const QString &name, const QByteArray &data, QString *errorMessage)
{
    if (!file_.isOpen()) {
        if (errorMessage) {
            *errorMessage = QObject::tr("ZIP output is not open.");
        }
        return false;
    }

    const QByteArray encodedName = name.toUtf8();
    Entry entry;
    entry.name = name;
    entry.crc = crc32(data);
    entry.size = static_cast<quint32>(data.size());
    entry.offset = static_cast<quint32>(file_.pos());

    constexpr quint16 flags = 0x0800;
    constexpr quint16 modTime = 0;
    constexpr quint16 modDate = 33;

    writeUInt32(0x04034b50);
    writeUInt16(20);
    writeUInt16(flags);
    writeUInt16(0);
    writeUInt16(modTime);
    writeUInt16(modDate);
    writeUInt32(entry.crc);
    writeUInt32(entry.size);
    writeUInt32(entry.size);
    writeUInt16(static_cast<quint16>(encodedName.size()));
    writeUInt16(0);
    file_.write(encodedName);
    file_.write(data);

    if (file_.error() != QFileDevice::NoError) {
        if (errorMessage) {
            *errorMessage = QObject::tr("Failed to write ZIP entry: %1").arg(file_.errorString());
        }
        return false;
    }

    entries_.append(entry);
    return true;
}

bool ZipWriter::close(QString *errorMessage)
{
    constexpr quint16 flags = 0x0800;
    constexpr quint16 modTime = 0;
    constexpr quint16 modDate = 33;
    const quint32 centralDirectoryOffset = static_cast<quint32>(file_.pos());

    for (const Entry &entry : entries_) {
        const QByteArray encodedName = entry.name.toUtf8();
        writeUInt32(0x02014b50);
        writeUInt16(20);
        writeUInt16(20);
        writeUInt16(flags);
        writeUInt16(0);
        writeUInt16(modTime);
        writeUInt16(modDate);
        writeUInt32(entry.crc);
        writeUInt32(entry.size);
        writeUInt32(entry.size);
        writeUInt16(static_cast<quint16>(encodedName.size()));
        writeUInt16(0);
        writeUInt16(0);
        writeUInt16(0);
        writeUInt16(0);
        writeUInt32(0);
        writeUInt32(entry.offset);
        file_.write(encodedName);
    }

    const quint32 centralDirectorySize = static_cast<quint32>(file_.pos()) - centralDirectoryOffset;
    writeUInt32(0x06054b50);
    writeUInt16(0);
    writeUInt16(0);
    writeUInt16(static_cast<quint16>(entries_.size()));
    writeUInt16(static_cast<quint16>(entries_.size()));
    writeUInt32(centralDirectorySize);
    writeUInt32(centralDirectoryOffset);
    writeUInt16(0);

    file_.close();
    if (file_.error() != QFileDevice::NoError) {
        if (errorMessage) {
            *errorMessage = QObject::tr("Failed to close ZIP output: %1").arg(file_.errorString());
        }
        return false;
    }
    return true;
}

quint32 ZipWriter::crc32(const QByteArray &data)
{
    static quint32 table[256] = {};
    static bool initialized = false;
    if (!initialized) {
        for (quint32 i = 0; i < 256; ++i) {
            quint32 value = i;
            for (int bit = 0; bit < 8; ++bit) {
                value = (value & 1) ? (0xedb88320U ^ (value >> 1)) : (value >> 1);
            }
            table[i] = value;
        }
        initialized = true;
    }

    quint32 crc = 0xffffffffU;
    for (const char byte : data) {
        crc = table[(crc ^ static_cast<uchar>(byte)) & 0xff] ^ (crc >> 8);
    }
    return crc ^ 0xffffffffU;
}

void ZipWriter::writeUInt16(quint16 value)
{
    char bytes[2] = {
        static_cast<char>(value & 0xff),
        static_cast<char>((value >> 8) & 0xff),
    };
    file_.write(bytes, 2);
}

void ZipWriter::writeUInt32(quint32 value)
{
    char bytes[4] = {
        static_cast<char>(value & 0xff),
        static_cast<char>((value >> 8) & 0xff),
        static_cast<char>((value >> 16) & 0xff),
        static_cast<char>((value >> 24) & 0xff),
    };
    file_.write(bytes, 4);
}
