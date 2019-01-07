#include "OrderEvent.h"

OrderEvent::OrderEvent(QObject *parent /*= Q_NULLPTR*/) : QObject(parent)
{
}

OrderEvent::~OrderEvent()
{
}

MySqlExThread::MySqlExThread(OrderCore* pCore, int runtype /*= 0*/) 
	: QThread()
	, m_runType(runtype)
	, m_pOrder(pCore)
{

}

MySqlExThread::~MySqlExThread()
{

}

void MySqlExThread::run()
{
	if (m_runType == MYSQL_ORDER_UPDATA_THREAD)
	{
		if (m_pOrder)
			m_pOrder->UpdataOrderTread();
	}
	if (m_runType == MYSQL_COMM_ADD_THREAD)
	{
		if (m_pOrder)
			m_pOrder->UpdataCommodityTread();
	}
}
