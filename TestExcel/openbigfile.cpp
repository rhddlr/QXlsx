#pragma once
#include <chrono>
#include <QDebug>

#include "xlsxdocument.h"

void openBigFile(bool isTest);

int openbigfile()
{
    openBigFile(true);
    return 0;
}

namespace {
class StopWatch
{
    using time_point = std::chrono::high_resolution_clock::time_point;

public:
    StopWatch() : _start(std::chrono::high_resolution_clock::now()) {}
    double elapsed()
    {
        const auto now = std::chrono::high_resolution_clock::now();
        const auto duration = std::chrono::duration<double, std::milli>(now - _start).count();
        _start = now;
        return duration;
    }

private:
    time_point _start;
};

void showCosts(double t, const QString& detail = QStringLiteral("operation"))
{
    qInfo() << QStringLiteral("%1: %2 (ms)").arg(detail).arg(t);
}

QXlsx::Format normalFormat()
{
    QXlsx::Format format;
    format.setNumberFormat("@");
    format.setFontBold(true);
    return format;
}

const int k_row = 5000;
const int k_column = 10;

void justReadData(const QString& file)
{
    StopWatch sw;
    {
        //![open big xlsx file, almost 200kb]
        QXlsx::Document bigXlsx(file);
        showCosts(sw.elapsed(), "open big xlsx file");

        if (!bigXlsx.isLoadPackage())
            qDebug() << QStringLiteral("[%1] open failed.").arg(file);

        //![move data to next row]
        const auto range = bigXlsx.dimension();
        for (int r = range.firstRow(); r < range.lastRow(); ++r)
            for (int c = range.firstColumn(); c < range.lastColumn(); ++c) {
                const auto data = bigXlsx.read(r, c);
                bigXlsx.write(r + 1, c, data);
            }

        showCosts(sw.elapsed(), QString("read data (%1 x %2)").arg(k_row).arg(k_column));

        if (!bigXlsx.saveAs(file)) // Default name is "Book1.xlsx"
        {
            qDebug() << QStringLiteral("[%1] failed to write excel.").arg(file);
            return;
        }
    }
    showCosts(sw.elapsed(), "release xlsx object");
}

void createFile(const QString& file)
{
    StopWatch sw;
    {
        //![Create a xlsx file]
        QXlsx::Document xlsx;

        // current sheet is Sheet1(default sheet)
        for (int i = 1; i < k_row; ++i) {
            for (int j = 1; j < k_column; ++j) {
                xlsx.write(i, j, QString("row %1 column %2").arg(i).arg(j));
            }
        }

        showCosts(sw.elapsed(), QString("write (%1 x %2) data").arg(k_row).arg(k_column));

        if (!xlsx.saveAs(file)) {
            qDebug() << QStringLiteral("[%1] failed to write excel.").arg(file);
            return;
        }
        showCosts(sw.elapsed(), QString("save (%1 x %2) data").arg(k_row).arg(k_column));
    }
    showCosts(sw.elapsed(), "release xlsx");
}

void moveDataToNextRow(const QString& file, bool withFormat = false)
{
    StopWatch sw;
    {
        //![open big xlsx file, almost 200kb]
        QXlsx::Document bigXlsx(file);
        showCosts(sw.elapsed(), "open big xlsx file");

        if (!bigXlsx.isLoadPackage())
            qDebug() << QStringLiteral("[%1] open failed.").arg(file);

        //![move data to next row]
        const auto range = bigXlsx.dimension();
        if (withFormat) {
            for (int r = range.firstRow(); r < range.lastRow(); ++r)
                for (int c = range.firstColumn(); c < range.lastColumn(); ++c) {
                    const auto data = bigXlsx.read(r, c);
                    bigXlsx.write(r + 1, c, data, normalFormat());
                }
        } else {
            for (int r = range.firstRow(); r < range.lastRow(); ++r)
                for (int c = range.firstColumn(); c < range.lastColumn(); ++c) {
                    const auto data = bigXlsx.read(r, c);
                    bigXlsx.write(r + 1, c, data);
                }
        }
        showCosts(sw.elapsed(), "rewrite data to next row");

        //![write data at the first row]
        for (int c = range.firstColumn(); c < range.lastColumn(); ++c)
            bigXlsx.write(1, c, c);
        showCosts(sw.elapsed(), QString("write one row (1 x %1) data").arg(k_column));

        if (!bigXlsx.saveAs(file)) // Default name is "Book1.xlsx"
        {
            qDebug() << QStringLiteral("[%1] failed to write excel.").arg(file);
            return;
        }
        showCosts(sw.elapsed(), QString("save (%1 x %2) data").arg(k_row).arg(k_column));
    }
    showCosts(sw.elapsed(), "release xlsx object");
}

} // namespace

void openBigFile(bool isTest)
{
    if (!isTest)
        return;
    const QString file{"openbigfile.xlsx"};
    qDebug() << "----------create a xlsx file--------------";
    createFile(file);
    qDebug() << "----------just read data--------------";
    justReadData(file);
    qDebug() << "----------read and write without format--------------";
    moveDataToNextRow(file, false);
    qDebug() << "----------read and write with format--------------";
    moveDataToNextRow(file, true);
}
