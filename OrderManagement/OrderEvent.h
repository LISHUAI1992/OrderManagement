#pragma once

#include <QObject>
#include "OrderDefine.h"
#include "OrderCore.h"

class OrderCore;

class OrderEvent : public QObject
{
	Q_OBJECT

public:
	OrderEvent(QObject *parent = Q_NULLPTR);
	~OrderEvent();

	void SetResult(const QString &result)
	{
		emit ExeSqlResult(result);
	}

signals:
	void ExeSqlResult(const QString &result);

};

class MySqlExThread : public QThread
{
public:
	MySqlExThread(OrderCore* pCore, int runtype = 0);
	virtual ~MySqlExThread();

	void run();

	void SetRunType(int runtype) { m_runType = runtype; }

private:
	int        m_runType;
	OrderCore* m_pOrder;
};

