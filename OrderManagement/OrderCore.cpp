#include "OrderCore.h"

OrderCore::OrderCore(OrderEvent* event, QObject *parent /*= Q_NULLPTR */) : QObject(parent)
	, m_mySqlDB(NULL)
	, m_pExcelAxObjet(NULL)
	, m_pSqlThread(NULL)
	, m_pOrderEvent(event)
{
	m_pExcelAxObjet = new QAxObject("Excel.Application");
	m_mySqlDB = new QSqlDatabase(QSqlDatabase::addDatabase("QMYSQL"));
	m_pSqlThread = new MySqlExThread(this);
}

OrderCore::~OrderCore()
{
	//m_pSqlThread->exit();
	SAFEDELETE(m_pSqlThread);
	SAFEDELETE(m_pExcelAxObjet);
	SAFEDELETE(m_mySqlDB);
}

void OrderCore::CastVariant2ListListVariant(const QVariant &var, ExcelList &res)
{
	QVariantList varRows = var.toList();
	if (varRows.isEmpty())
	{
		return;
	}
	const int rowCount = varRows.size();
	QVariantList rowData;
	for (int i = 0; i < rowCount; ++i)
	{
		rowData = varRows[i].toList();
		res.push_back(rowData);
	}
}

bool OrderCore::OpenMySqlDB(const QString& hostName, const QString& dataBaseName,
	const QString& userName, const QString& passWord)
{
	if (!m_mySqlDB)
		return false;
	m_mySqlDB->setHostName(hostName);
	m_mySqlDB->setDatabaseName(dataBaseName);
	m_mySqlDB->setUserName(userName);
	m_mySqlDB->setPassword(passWord);
	if (!m_mySqlDB->open())
	{
		QMessageBox::critical(0, TU("错误"), TU("无法创建数据库连接"), QMessageBox::Cancel);
		QSqlError err = m_mySqlDB->lastError();
		qDebug() << "connect DB Err:" << err.text();
		return false;
	}
	return true;
}

void OrderCore::CloseMySqlDB()
{
	if (m_mySqlDB)
		m_mySqlDB->close();
}

void OrderCore::UpdataOrderTread()
{
	if (!m_mySqlDB)
		return;
	QSqlQuery query(*m_mySqlDB);
	QString sql = "";
	QString restul;
	bool bRet = false;
	for (size_t i = 1; i < m_excelListData.size(); i++)
	{
		sql = "";
		restul = "";
		if (IsExistOrder(m_excelListData.at(i).at(0).toString()))
			BuildUpdataOrderSql(m_excelListData.at(i), sql);
		else
			BuildAddOrderSql(m_excelListData.at(i), sql);
		if (!sql.isEmpty())
			bRet = query.exec(sql);
		if (bRet)
		{
			restul = QString(TU("第%1条数据导入成功")).arg(i);
		}
		else
		{
			restul = QString(TU("第%1条数据导入失败")).arg(i);
		}
		m_pOrderEvent->SetResult(restul);
	}
}

void OrderCore::UpdataCommodityTread()
{
	if (!m_mySqlDB)
		return;
	QSqlQuery query(*m_mySqlDB);
	QString sql = "";
	QString restul;
	bool bRet = false;
	for (size_t i = 1; i < m_excelListData.size(); i++)
	{
		sql = "";
		restul = "";
		BuildAddCommSql(m_excelListData.at(i), sql);
		if (!sql.isEmpty())
			bRet = query.exec(sql);
		if (bRet)
		{
			restul = QString(TU("第%1条数据导入成功")).arg(i);
		}
		else
		{
			restul = QString(TU("第%1条数据导入失败")).arg(i);
		}
		m_pOrderEvent->SetResult(restul);
	}
}

bool OrderCore::UpdataCommodity(const ExcelList& data)
{
	m_excelListData = data;
	m_pSqlThread->SetRunType(MYSQL_COMM_ADD_THREAD);
	m_pSqlThread->start();
	return true;
}

bool OrderCore::UpdataOrder(const ExcelList& data)
{
	m_excelListData = data;
	m_pSqlThread->SetRunType(MYSQL_ORDER_UPDATA_THREAD);
	m_pSqlThread->start();
	return true;
}

void OrderCore::BuildAddCommSql(const ExcelRow& rowData, QString& sql)
{

	//"INSERT INTO commodity (order_indexcode, title_txt, price, number, othersys_indexcode, attribute, package_info, remarks, status, merchant_code)"
	//"VALUES ('%s', '%s', %d, %d, '%s', '%s', '%s', '%s', '%s', '%s')
	QString tempSql = "INSERT INTO commodity (";
	QString tempValues = ")VALUES (";
	if (rowData.size() < 10)
		return;
	tempSql += "order_indexcode";
	tempValues += QString("'%1'").arg(rowData.at(0).toString());
	AddValue(tempSql, tempValues, rowData.at(1), "title_txt");
	AddValue(tempSql, tempValues, rowData.at(2), "price");
	AddValue(tempSql, tempValues, rowData.at(3), "number");
	AddValue(tempSql, tempValues, rowData.at(4), "othersys_indexcode");
	AddValue(tempSql, tempValues, rowData.at(5), "attribute");
	AddValue(tempSql, tempValues, rowData.at(6), "package_info");
	AddValue(tempSql, tempValues, rowData.at(7), "remarks");
	AddValue(tempSql, tempValues, rowData.at(8), "status");
	AddValue(tempSql, tempValues, rowData.at(9), "merchant_code");
	tempSql += ", createtime";
	tempValues += ", NOW()";

	sql = tempSql + tempValues + ")";
}

void OrderCore::BuildAddOrderSql(const ExcelRow& rowData, QString& sql)
{
	QString tempSql = "INSERT INTO `order` (";
	QString tempValues = ")VALUES (";
	if (rowData.size() < 10)
		return;
	tempSql += "order_indexcode";
	tempValues += QString("'%1'").arg(rowData.at(0).toString());
	AddValue(tempSql, tempValues, rowData.at(1), "buyer_name");
	AddValue(tempSql, tempValues, rowData.at(2), "b_pay_account");
	AddValue(tempSql, tempValues, rowData.at(3), "b_pay_code");
	AddValue(tempSql, tempValues, rowData.at(4), "b_pay_details");

	AddValue(tempSql, tempValues, rowData.at(5), "b_pay_monay");
	AddValue(tempSql, tempValues, rowData.at(6), "b_pay_postage");
	AddValue(tempSql, tempValues, rowData.at(7), "b_pay_integral");
	AddValue(tempSql, tempValues, rowData.at(8), "b_total_monay");
	AddValue(tempSql, tempValues, rowData.at(9), "b_rebates_integral");
	AddValue(tempSql, tempValues, rowData.at(10), "b_realpay_monay");
	AddValue(tempSql, tempValues, rowData.at(11), "b_realpay_integral");

	AddValue(tempSql, tempValues, rowData.at(12), "b_order_state");
	AddValue(tempSql, tempValues, rowData.at(13), "b_note");
	AddValue(tempSql, tempValues, rowData.at(14), "b_recipient_name");
	AddValue(tempSql, tempValues, rowData.at(15), "b_recipient_adress");
	AddValue(tempSql, tempValues, rowData.at(16), "o_transport_type");
	AddValue(tempSql, tempValues, rowData.at(17), "b_recipient_phone");
	AddValue(tempSql, tempValues, rowData.at(18), "b_recipient_mphone");
	AddValue(tempSql, tempValues, rowData.at(19), "b_order_createtime");
	AddValue(tempSql, tempValues, rowData.at(20), "b_order_paytime");


	AddValue(tempSql, tempValues, rowData.at(21), "o_commodity_note");
	AddValue(tempSql, tempValues, rowData.at(22), "o_commodity_type");
	AddValue(tempSql, tempValues, rowData.at(23), "o_logistic_code");
	AddValue(tempSql, tempValues, rowData.at(24), "o_logistic_company");
	AddValue(tempSql, tempValues, rowData.at(25), "o_order_note");
	AddValue(tempSql, tempValues, rowData.at(26), "o_commodity_count");

	AddValue(tempSql, tempValues, rowData.at(27), "o_shop_id");
	AddValue(tempSql, tempValues, rowData.at(28), "o_shop_name");
	AddValue(tempSql, tempValues, rowData.at(29), "o_order_closereson");
	AddValue(tempSql, tempValues, rowData.at(30), "s_seller_fee");
	AddValue(tempSql, tempValues, rowData.at(31), "b_buyer_fee");
	AddValue(tempSql, tempValues, rowData.at(32), "o_invoice_info");
	AddValue(tempSql, tempValues, rowData.at(33), "o_phon_order");
	AddValue(tempSql, tempValues, rowData.at(34), "o_phaseorder_info");

	AddValue(tempSql, tempValues, rowData.at(35), "o_privilegeorder_id");
	AddValue(tempSql, tempValues, rowData.at(36), "o_contract_pic");
	AddValue(tempSql, tempValues, rowData.at(37), "o_order_receipts");
	AddValue(tempSql, tempValues, rowData.at(38), "o_order_paid");
	AddValue(tempSql, tempValues, rowData.at(39), "o_deposit_rank");
	AddValue(tempSql, tempValues, rowData.at(40), "o_modified_sku");

	AddValue(tempSql, tempValues, rowData.at(41), "o_modified_adress");
	AddValue(tempSql, tempValues, rowData.at(42), "o_abnormal_info");
	AddValue(tempSql, tempValues, rowData.at(43), "o_tmall_voucher");
	AddValue(tempSql, tempValues, rowData.at(44), "o_jifenbao_voucher");
	AddValue(tempSql, tempValues, rowData.at(45), "o_o2o_trading");
	AddValue(tempSql, tempValues, rowData.at(46), "o_trading_type");
	AddValue(tempSql, tempValues, rowData.at(47), "o_retailshop_name");
	AddValue(tempSql, tempValues, rowData.at(48), "o_retailshop_id");
	AddValue(tempSql, tempValues, rowData.at(49), "o_retaildelivery_name");
	AddValue(tempSql, tempValues, rowData.at(50), "o_retaildelivery_id");

	AddValue(tempSql, tempValues, rowData.at(51), "o_refund_account");
	AddValue(tempSql, tempValues, rowData.at(52), "o_appointment_shop");
	AddValue(tempSql, tempValues, rowData.at(53), "b_order_confirmtime");
	AddValue(tempSql, tempValues, rowData.at(54), "b_pay_confirmaccount");
	AddValue(tempSql, tempValues, rowData.at(55), "o_buyer_envelope");
	AddValue(tempSql, tempValues, rowData.at(56), "o_mainorder_indexcode");
	if (rowData.size() > 57)
	{
		AddValue(tempSql, tempValues, rowData.at(57), "o_ext1_info");
		AddValue(tempSql, tempValues, rowData.at(58), "o_ext2_info");
	}

	tempSql += ", createtime";
	tempValues += ", NOW()";

	tempSql += ", versions";
	tempValues += ", 0";

	sql = tempSql + tempValues + ")";
}

void OrderCore::BuildUpdataOrderSql(const ExcelRow& rowData, QString& sql)
{
	QString tempSql = "UPDATE `order` SET ";
	QString tempValues = QString(" WHERE `order_indexcode` = Cast( '%1' AS BINARY ( 18 ) );").arg(rowData.at(0).toString());
	if (rowData.size() < 10)
		return;
	tempSql += QString("`buyer_name` = '%1'").arg(rowData.at(1).toString());
	UpdataValue(tempSql, rowData.at(2), "b_pay_account");
	UpdataValue(tempSql, rowData.at(3), "b_pay_code");
	UpdataValue(tempSql, rowData.at(4), "b_pay_details");

	UpdataValue(tempSql, rowData.at(5), "b_pay_monay");
	UpdataValue(tempSql, rowData.at(6), "b_pay_postage");
	UpdataValue(tempSql, rowData.at(7), "b_pay_integral");
	UpdataValue(tempSql, rowData.at(8), "b_total_monay");
	UpdataValue(tempSql, rowData.at(9), "b_rebates_integral");
	UpdataValue(tempSql, rowData.at(10), "b_realpay_monay");
	UpdataValue(tempSql, rowData.at(11), "b_realpay_integral");

	UpdataValue(tempSql, rowData.at(12), "b_order_state");
	UpdataValue(tempSql, rowData.at(13), "b_note");
	UpdataValue(tempSql, rowData.at(14), "b_recipient_name");
	UpdataValue(tempSql, rowData.at(15), "b_recipient_adress");
	UpdataValue(tempSql, rowData.at(16), "o_transport_type");
	UpdataValue(tempSql, rowData.at(17), "b_recipient_phone");
	UpdataValue(tempSql, rowData.at(18), "b_recipient_mphone");
	UpdataValue(tempSql, rowData.at(19), "b_order_createtime");
	UpdataValue(tempSql, rowData.at(20), "b_order_paytime");


	UpdataValue(tempSql, rowData.at(21), "o_commodity_note");
	UpdataValue(tempSql, rowData.at(22), "o_commodity_type");
	UpdataValue(tempSql, rowData.at(23), "o_logistic_code");
	UpdataValue(tempSql, rowData.at(24), "o_logistic_company");
	UpdataValue(tempSql, rowData.at(25), "o_order_note");
	UpdataValue(tempSql, rowData.at(26), "o_commodity_count");

	UpdataValue(tempSql, rowData.at(27), "o_shop_id");
	UpdataValue(tempSql, rowData.at(28), "o_shop_name");
	UpdataValue(tempSql, rowData.at(29), "o_order_closereson");
	UpdataValue(tempSql, rowData.at(30), "s_seller_fee");
	UpdataValue(tempSql, rowData.at(31), "b_buyer_fee");
	UpdataValue(tempSql, rowData.at(32), "o_invoice_info");
	UpdataValue(tempSql, rowData.at(33), "o_phon_order");
	UpdataValue(tempSql, rowData.at(34), "o_phaseorder_info");

	UpdataValue(tempSql, rowData.at(35), "o_privilegeorder_id");
	UpdataValue(tempSql, rowData.at(36), "o_contract_pic");
	UpdataValue(tempSql, rowData.at(37), "o_order_receipts");
	UpdataValue(tempSql, rowData.at(38), "o_order_paid");
	UpdataValue(tempSql, rowData.at(39), "o_deposit_rank");
	UpdataValue(tempSql, rowData.at(40), "o_modified_sku");

	UpdataValue(tempSql, rowData.at(41), "o_modified_adress");
	UpdataValue(tempSql, rowData.at(42), "o_abnormal_info");
	UpdataValue(tempSql, rowData.at(43), "o_tmall_voucher");
	UpdataValue(tempSql, rowData.at(44), "o_jifenbao_voucher");
	UpdataValue(tempSql, rowData.at(45), "o_o2o_trading");
	UpdataValue(tempSql, rowData.at(46), "o_trading_type");
	UpdataValue(tempSql, rowData.at(47), "o_retailshop_name");
	UpdataValue(tempSql, rowData.at(48), "o_retailshop_id");
	UpdataValue(tempSql, rowData.at(49), "o_retaildelivery_name");
	UpdataValue(tempSql, rowData.at(50), "o_retaildelivery_id");

	UpdataValue(tempSql, rowData.at(51), "o_refund_account");
	UpdataValue(tempSql, rowData.at(52), "o_appointment_shop");
	UpdataValue(tempSql, rowData.at(53), "b_order_confirmtime");
	UpdataValue(tempSql, rowData.at(54), "b_pay_confirmaccount");
	UpdataValue(tempSql, rowData.at(55), "o_buyer_envelope");
	UpdataValue(tempSql, rowData.at(56), "o_mainorder_indexcode");
	if (rowData.size() > 57)
	{
		UpdataValue(tempSql, rowData.at(57), "o_ext1_info");
		UpdataValue(tempSql, rowData.at(58), "o_ext2_info");
	}

	tempSql += ", `updatatime` = NOW()";
	tempSql += ", `versions` = `versions`+1";

	sql = tempSql + tempValues;
}

bool OrderCore::IsExistOrder(const QString& indexCode)
{
	QSqlQuery query(*m_mySqlDB);
	QString sql = QString("SELECT * FROM `order` WHERE order_indexcode = %1").arg(indexCode);
	query.exec(sql);
	return query.next();
}

void OrderCore::AddValue(QString& sql, QString& values, const QVariant& var, const QString& marke)
{
	if (!var.isValid() || var.isNull())
	{
		sql += ", " + marke;
		values += ", NULL";
		return;
	}

	switch (var.type())
	{
	case QVariant::String:
	{
		QString tempData = var.toString();
		if (tempData != "null" && !tempData.isEmpty())
		{
			tempData.replace("'", "");
			sql += ", " + marke;
			values += QString(", '%1'").arg(tempData);
		}
		break;
	}
	case QVariant::Int:
	{
		int tempData = var.toInt();
		sql += ", " + marke;
		values += QString(", %1").arg(tempData);
		break;
	}
	case QVariant::Double:
	{
		double tempData = var.toDouble();
		sql += ", " + marke;
		values += QString(", %1").arg(tempData);
		break;
	}
	case QVariant::DateTime:
	{
		QString tempData = (var.toDateTime()).toString("yyyy-MM-dd hh:mm:ss");
		sql += ", " + marke;
		values += QString(", '%1'").arg(tempData);
		break;
	}
	default:
		break;
	}
}

void OrderCore::UpdataValue(QString& sql, const QVariant& var, const QString& marke)
{
	if (!var.isValid() || var.isNull())
		return;
	switch (var.type())
	{
	case QVariant::String:
	{
		QString tempData = var.toString();
		if (tempData != "null" && !tempData.isEmpty())
		{
			sql += QString(",`%1` = '%2'").arg(marke).arg(tempData);
		}
		break;
	}
	case QVariant::Int:
	{
		int tempData = var.toInt();
		sql += QString(",`%1` = %2").arg(marke).arg(tempData);
		break;
	}
	case QVariant::Double:
	{
		double tempData = var.toDouble();
		sql += QString(",`%1` = %2").arg(marke).arg(tempData);
		break;
	}
	case QVariant::DateTime:
	{
		QString tempData = (var.toDateTime()).toString("yyyy-MM-dd hh:mm:ss");
		sql += QString(",`%1` = '%2'").arg(marke).arg(tempData);
		break;
	}
	default:
		break;
	}
}

void OrderCore::ReadExcelData(const QString& excelFilePath, ExcelList& excelList)
{
	if (m_pExcelAxObjet == NULL)
		return;

	m_pExcelAxObjet->setProperty("Visible", false);
	QAxObject* workbooks = m_pExcelAxObjet->querySubObject("WorkBooks");
	if (workbooks == NULL)
		return;
	workbooks->dynamicCall("Open (const QString&)", excelFilePath); //打开文件
	QAxObject* workbook = m_pExcelAxObjet->querySubObject("ActiveWorkBook");
	if (workbook == NULL)
		return;
	QAxObject* worksheets = workbook->querySubObject("WorkSheets");
	if (worksheets == NULL)
		return;
	QAxObject* worksheet = workbook->querySubObject("Worksheets(int)", 1); //打开第一个
	if (worksheet == NULL)
		return;
	QAxObject* usedrange = worksheet->querySubObject("UsedRange");
	if (usedrange == NULL)
		return;
	QVariant var;
	var = usedrange->dynamicCall("Value");
	CastVariant2ListListVariant(var, excelList);
}
