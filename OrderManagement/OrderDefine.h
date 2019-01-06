#pragma once
#include <QVariant>

#define TU(s) QString::fromLocal8Bit(s)

typedef QList<QVariant>   ExcelRow;

typedef QList<QList<QVariant>>  ExcelList;