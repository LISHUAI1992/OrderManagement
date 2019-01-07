#pragma once

#include <QObject>
#include "OrderDefine.h"
#include "OrderEvent.h"

class MySqlExThread;
class OrderEvent;

class OrderCore : public QObject
{
	Q_OBJECT

public:
	OrderCore(OrderEvent* event, QObject *parent = Q_NULLPTR );
	~OrderCore();

	void ReadExcelData(const QString& excelFilePath, ExcelList& excelList);

	bool UpdataCommodity(const ExcelList& data);

	bool UpdataOrder(const ExcelList& data);

	bool OpenMySqlDB(const QString& hostName, const QString& dataBaseName,
		const QString& userName, const QString& passWord);
	void CloseMySqlDB();

	void UpdataOrderTread();

	void UpdataCommodityTread();

protected:

	void CastVariant2ListListVariant(const QVariant &var, ExcelList &res);

	void BuildAddCommSql(const ExcelRow& rowData, QString& sql);

	void BuildAddOrderSql(const ExcelRow& rowData, QString& sql);

	void BuildUpdataOrderSql(const ExcelRow& rowData, QString& sql);

	bool IsExistOrder(const QString& indexCode);

	void AddValue(QString& sql, QString& values, const QVariant& var, const QString& marke);

	void UpdataValue(QString& sql, const QVariant& var, const QString& marke);

private:
	QAxObject*      m_pExcelAxObjet;
	QSqlDatabase*   m_mySqlDB;
	ExcelList       m_excelListData;
	MySqlExThread*  m_pSqlThread;
	OrderEvent*     m_pOrderEvent;
};
