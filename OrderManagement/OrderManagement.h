#pragma once

#include <QtWidgets/QMainWindow>
#include "ui_OrderManagement.h"
#include "MySQLInfo.h"
#include "OrderCore.h"
#include "OrderEvent.h"
#include "ImportWidget.h"

class OrderEvent;

class OrderManagement : public QMainWindow
{
	Q_OBJECT

public:
	OrderManagement(QWidget *parent = Q_NULLPTR);
	virtual ~OrderManagement();

protected:


private slots:
	void  openExcelFile();
	void  openDBInfoWidget();
	void  on_openMySql(const QString &host, const QString &user, const QString &password);

private:
	Ui::OrderManagementClass ui;
	MySQLInfo* m_mySqlInfo;
	OrderCore* m_pCore;
	OrderEvent* m_pEvent;
	ImportWidget* m_pImportData;
};
