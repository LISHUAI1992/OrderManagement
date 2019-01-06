#pragma once

#include <QWidget>
#include "ui_MySQLInfo.h"
#include "OrderDefine.h"

class MySQLInfo : public QWidget
{
	Q_OBJECT

public:
	MySQLInfo(QWidget *parent = Q_NULLPTR);
	~MySQLInfo();

signals:
	void openMysql(const QString &host, const QString &user, const QString &password);


private slots:
	void  on_saveBtn_clicked();

private:
	Ui::MySQLInfo ui;
	QString m_hostName;
	QString m_userName;
	QString m_passWord;
};
