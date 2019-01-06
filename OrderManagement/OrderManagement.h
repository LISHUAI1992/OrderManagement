#pragma once

#include <QtWidgets/QMainWindow>
#include <ActiveQt/QAxObject>
#include <QFileDialog>
#include <QDebug>
#include <QMessageBox>
#include <QSqlDatabase>
#include <QSqlQuery>
#include "ui_OrderManagement.h"
#include "MySQLInfo.h"


class OrderManagement : public QMainWindow
{
	Q_OBJECT

public:
	OrderManagement(QWidget *parent = Q_NULLPTR);
	virtual ~OrderManagement();

protected:
	void readExcelData(const QString& excelFilePath, ExcelList& excelList);

	void castVariant2ListListVariant(const QVariant &var, ExcelList &res);

	bool openMySqlDB(const QString& hostName, const QString& dataBaseName,
		const QString& userName, const QString& passWord);
	void closeMySqlDB();

	bool updataCommodity(const ExcelList& data);

	bool updataOrder(const ExcelList& data);

	void buildAddCommSql(const ExcelRow& rowData, QString& sql);

	void buildAddOrderSql(const ExcelRow& rowData, QString& sql);

	void buildUpdataOrderSql(const ExcelRow& rowData, QString& sql);

	bool isExistOrder(const QString& indexCode);

	void addValue(QString& sql, QString& values, const QVariant& var, const QString& marke);

	void updataValue(QString& sql, const QVariant& var, const QString& marke);

private slots:
	void  openExcelFile();
	void  openDBInfoWidget();
	void  on_openMySql(const QString &host, const QString &user, const QString &password);

private:
	Ui::OrderManagementClass ui;
	QAxObject * m_pExcelAxObjet;
	QSqlDatabase* m_mySqlDB;
	MySQLInfo* m_mySqlInfo;
};
