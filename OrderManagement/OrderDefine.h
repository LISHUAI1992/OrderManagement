#pragma once
#include <QVariant>
#include <ActiveQt/QAxObject>
#include <QFileDialog>
#include <QDebug>
#include <QMessageBox>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QDateTime>
#include <QThread>

#define TU(s) QString::fromLocal8Bit(s)

#define SAFENEW   new (std::nothrow) 

#define SAFEDELETE(s) {if(s != NULL){delete s;s = NULL;}}

//#define MYSQL_ORDER_ADD_THREAD      0
#define MYSQL_ORDER_UPDATA_THREAD   0
#define MYSQL_COMM_ADD_THREAD       1

typedef QList<QVariant>   ExcelRow;

typedef QList<QList<QVariant>>  ExcelList;