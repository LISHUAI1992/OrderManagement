#include "OrderManagement.h"
#include <QSqlError>
#include <QDateTime>

OrderManagement::OrderManagement(QWidget *parent)
	: QMainWindow(parent)
	, m_pExcelAxObjet(NULL)
	, m_mySqlInfo(NULL)
	, m_mySqlDB(NULL)
{
	ui.setupUi(this);

	setWindowTitle(TU("订单管理系统"));
	m_mySqlInfo = new MySQLInfo();
	m_mySqlInfo->hide();
	m_pExcelAxObjet = new QAxObject("Excel.Application");

	m_mySqlDB = new QSqlDatabase(QSqlDatabase::addDatabase("QMYSQL"));

	ui.actionupdatafile->setText(TU("上传文件"));
	ui.actionsetDB->setText(TU("设置数据库信息"));
	
	connect(ui.actionupdatafile, SIGNAL(triggered()),
		this, SLOT(openExcelFile()));
	connect(ui.actionsetDB, SIGNAL(triggered()),
		this, SLOT(openDBInfoWidget()));
	connect(m_mySqlInfo, SIGNAL(openMysql(const QString &, const QString &, const QString &)),
		this, SLOT(on_openMySql(const QString &, const QString &, const QString &)));
	openMySqlDB("132.232.101.227", "shop", "myuser", "Hik19920623#123");
}

OrderManagement::~OrderManagement()
{
	if (m_pExcelAxObjet == NULL)
		delete m_pExcelAxObjet;
	if (!m_mySqlDB)
		delete m_mySqlDB;
	if (!m_mySqlInfo)
		delete m_mySqlInfo;
}


void OrderManagement::castVariant2ListListVariant(const QVariant &var, ExcelList &res)
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

bool OrderManagement::openMySqlDB(const QString& hostName, const QString& dataBaseName, 
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
		QMessageBox::critical(0, TU("错误"),TU("无法创建数据库连接"), QMessageBox::Cancel);
		QSqlError err = m_mySqlDB->lastError();
		qDebug() << "connect DB Err:" << err.text();
		return false;
	}	
	return true;
}

void OrderManagement::closeMySqlDB()
{
	if (m_mySqlDB)
		m_mySqlDB->close();
}

bool OrderManagement::updataCommodity(const ExcelList& data)
{
	if (!m_mySqlDB)
		return false;
	QSqlQuery query(*m_mySqlDB);
	QString sql = "";
	for (size_t i = 1;i < data.size(); i++)
	{
		sql = "";
		buildAddCommSql(data.at(i), sql);
		if (!sql.isEmpty())
			query.exec(sql);
	}
	return true;
}

bool OrderManagement::updataOrder(const ExcelList& data)
{
	if (!m_mySqlDB)
		return false;
	QSqlQuery query(*m_mySqlDB);
	QString sql = "";
	for (size_t i = 1; i < data.size(); i++)
	{
		sql = "";
		if (isExistOrder(data.at(i).at(0).toString()))
			buildUpdataOrderSql(data.at(i), sql);
		else
			buildAddOrderSql(data.at(i), sql);
		if (!sql.isEmpty())
			query.exec(sql);
	}
	return true;
}

void OrderManagement::buildAddCommSql(const ExcelRow& rowData, QString& sql)
{

	//"INSERT INTO commodity (order_indexcode, title_txt, price, number, othersys_indexcode, attribute, package_info, remarks, status, merchant_code)"
    //"VALUES ('%s', '%s', %d, %d, '%s', '%s', '%s', '%s', '%s', '%s')
	QString tempSql = "INSERT INTO commodity (";
	QString tempValues = ")VALUES (";
	if (rowData.size() < 10)
		return;
	tempSql += "order_indexcode";
	tempValues += QString("'%1'").arg(rowData.at(0).toString());
	addValue(tempSql, tempValues, rowData.at(1), "title_txt");
	addValue(tempSql, tempValues, rowData.at(2), "price");
	addValue(tempSql, tempValues, rowData.at(3), "number");
	addValue(tempSql, tempValues, rowData.at(4), "othersys_indexcode");
	addValue(tempSql, tempValues, rowData.at(5), "attribute");
	addValue(tempSql, tempValues, rowData.at(6), "package_info");
	addValue(tempSql, tempValues, rowData.at(7), "remarks");
	addValue(tempSql, tempValues, rowData.at(8), "status");
	addValue(tempSql, tempValues, rowData.at(9), "merchant_code");
	tempSql += ", createtime";
	tempValues += ", NOW()";

	sql = tempSql + tempValues + ")";
}

void OrderManagement::buildAddOrderSql(const ExcelRow& rowData, QString& sql)
{
	QString tempSql = "INSERT INTO `order` (";
	QString tempValues = ")VALUES (";
	if (rowData.size() < 10)
		return;
	tempSql += "order_indexcode";
	tempValues += QString("'%1'").arg(rowData.at(0).toString());
	addValue(tempSql, tempValues, rowData.at(1), "buyer_name");
	addValue(tempSql, tempValues, rowData.at(2), "b_pay_account");
	addValue(tempSql, tempValues, rowData.at(3), "b_pay_code");
	addValue(tempSql, tempValues, rowData.at(4), "b_pay_details");

	addValue(tempSql, tempValues, rowData.at(5), "b_pay_monay");
	addValue(tempSql, tempValues, rowData.at(6), "b_pay_postage");
	addValue(tempSql, tempValues, rowData.at(7), "b_pay_integral");
	addValue(tempSql, tempValues, rowData.at(8), "b_total_monay");
	addValue(tempSql, tempValues, rowData.at(9), "b_rebates_integral");
	addValue(tempSql, tempValues, rowData.at(10), "b_realpay_monay");
	addValue(tempSql, tempValues, rowData.at(11), "b_realpay_integral");

	addValue(tempSql, tempValues, rowData.at(12), "b_order_state");
	addValue(tempSql, tempValues, rowData.at(13), "b_note");
	addValue(tempSql, tempValues, rowData.at(14), "b_recipient_name");
	addValue(tempSql, tempValues, rowData.at(15), "b_recipient_adress");
	addValue(tempSql, tempValues, rowData.at(16), "o_transport_type");
	addValue(tempSql, tempValues, rowData.at(17), "b_recipient_phone");
	addValue(tempSql, tempValues, rowData.at(18), "b_recipient_mphone");
	addValue(tempSql, tempValues, rowData.at(19), "b_order_createtime");
	addValue(tempSql, tempValues, rowData.at(20), "b_order_paytime");


	addValue(tempSql, tempValues, rowData.at(21), "o_commodity_note");
	addValue(tempSql, tempValues, rowData.at(22), "o_commodity_type");
	addValue(tempSql, tempValues, rowData.at(23), "o_logistic_code");
	addValue(tempSql, tempValues, rowData.at(24), "o_logistic_company");
	addValue(tempSql, tempValues, rowData.at(25), "o_order_note");
	addValue(tempSql, tempValues, rowData.at(26), "o_commodity_count");

	addValue(tempSql, tempValues, rowData.at(27), "o_shop_id");
	addValue(tempSql, tempValues, rowData.at(28), "o_shop_name");
	addValue(tempSql, tempValues, rowData.at(29), "o_order_closereson");
	addValue(tempSql, tempValues, rowData.at(30), "s_seller_fee");
	addValue(tempSql, tempValues, rowData.at(31), "b_buyer_fee");
	addValue(tempSql, tempValues, rowData.at(32), "o_invoice_info");
	addValue(tempSql, tempValues, rowData.at(33), "o_phon_order");
	addValue(tempSql, tempValues, rowData.at(34), "o_phaseorder_info");

	addValue(tempSql, tempValues, rowData.at(35), "o_privilegeorder_id");
	addValue(tempSql, tempValues, rowData.at(36), "o_contract_pic");
	addValue(tempSql, tempValues, rowData.at(37), "o_order_receipts");
	addValue(tempSql, tempValues, rowData.at(38), "o_order_paid");
	addValue(tempSql, tempValues, rowData.at(39), "o_deposit_rank");
	addValue(tempSql, tempValues, rowData.at(40), "o_modified_sku");

	addValue(tempSql, tempValues, rowData.at(41), "o_modified_adress");
	addValue(tempSql, tempValues, rowData.at(42), "o_abnormal_info");
	addValue(tempSql, tempValues, rowData.at(43), "o_tmall_voucher");
	addValue(tempSql, tempValues, rowData.at(44), "o_jifenbao_voucher");
	addValue(tempSql, tempValues, rowData.at(45), "o_o2o_trading");
	addValue(tempSql, tempValues, rowData.at(46), "o_trading_type");
	addValue(tempSql, tempValues, rowData.at(47), "o_retailshop_name");
	addValue(tempSql, tempValues, rowData.at(48), "o_retailshop_id");
	addValue(tempSql, tempValues, rowData.at(49), "o_retaildelivery_name");
	addValue(tempSql, tempValues, rowData.at(50), "o_retaildelivery_id");

	addValue(tempSql, tempValues, rowData.at(51), "o_refund_account");
	addValue(tempSql, tempValues, rowData.at(52), "o_appointment_shop");
	addValue(tempSql, tempValues, rowData.at(53), "b_order_confirmtime");
	addValue(tempSql, tempValues, rowData.at(54), "b_pay_confirmaccount");
	addValue(tempSql, tempValues, rowData.at(55), "o_buyer_envelope");
	addValue(tempSql, tempValues, rowData.at(56), "o_mainorder_indexcode");
	addValue(tempSql, tempValues, rowData.at(57), "o_ext1_info");
	addValue(tempSql, tempValues, rowData.at(58), "o_ext2_info");

	tempSql += ", createtime";
	tempValues += ", NOW()";

	tempSql += ", versions";
	tempValues += ", 0";

	sql = tempSql + tempValues + ")";
}

void OrderManagement::buildUpdataOrderSql(const ExcelRow& rowData, QString& sql)
{
	QString tempSql = "UPDATE `order` SET ";
	QString tempValues = QString(" WHERE `order_indexcode` = Cast( '%1' AS BINARY ( 18 ) );").arg(rowData.at(0).toString());
	if (rowData.size() < 10)
		return;
	tempSql += QString("`buyer_name` = '%1'").arg(rowData.at(1).toString());
	updataValue(tempSql, rowData.at(2), "b_pay_account");
	updataValue(tempSql, rowData.at(3), "b_pay_code");
	updataValue(tempSql, rowData.at(4), "b_pay_details");

	updataValue(tempSql, rowData.at(5), "b_pay_monay");
	updataValue(tempSql, rowData.at(6), "b_pay_postage");
	updataValue(tempSql, rowData.at(7), "b_pay_integral");
	updataValue(tempSql, rowData.at(8), "b_total_monay");
	updataValue(tempSql, rowData.at(9), "b_rebates_integral");
	updataValue(tempSql, rowData.at(10), "b_realpay_monay");
	updataValue(tempSql, rowData.at(11), "b_realpay_integral");

	updataValue(tempSql, rowData.at(12), "b_order_state");
	updataValue(tempSql, rowData.at(13), "b_note");
	updataValue(tempSql, rowData.at(14), "b_recipient_name");
	updataValue(tempSql, rowData.at(15), "b_recipient_adress");
	updataValue(tempSql, rowData.at(16), "o_transport_type");
	updataValue(tempSql, rowData.at(17), "b_recipient_phone");
	updataValue(tempSql, rowData.at(18), "b_recipient_mphone");
	updataValue(tempSql, rowData.at(19), "b_order_createtime");
	updataValue(tempSql, rowData.at(20), "b_order_paytime");


	updataValue(tempSql, rowData.at(21), "o_commodity_note");
	updataValue(tempSql, rowData.at(22), "o_commodity_type");
	updataValue(tempSql, rowData.at(23), "o_logistic_code");
	updataValue(tempSql, rowData.at(24), "o_logistic_company");
	updataValue(tempSql, rowData.at(25), "o_order_note");
	updataValue(tempSql, rowData.at(26), "o_commodity_count");

	updataValue(tempSql, rowData.at(27), "o_shop_id");
	updataValue(tempSql, rowData.at(28), "o_shop_name");
	updataValue(tempSql, rowData.at(29), "o_order_closereson");
	updataValue(tempSql, rowData.at(30), "s_seller_fee");
	updataValue(tempSql, rowData.at(31), "b_buyer_fee");
	updataValue(tempSql, rowData.at(32), "o_invoice_info");
	updataValue(tempSql, rowData.at(33), "o_phon_order");
	updataValue(tempSql, rowData.at(34), "o_phaseorder_info");

	updataValue(tempSql, rowData.at(35), "o_privilegeorder_id");
	updataValue(tempSql, rowData.at(36), "o_contract_pic");
	updataValue(tempSql, rowData.at(37), "o_order_receipts");
	updataValue(tempSql, rowData.at(38), "o_order_paid");
	updataValue(tempSql, rowData.at(39), "o_deposit_rank");
	updataValue(tempSql, rowData.at(40), "o_modified_sku");

	updataValue(tempSql, rowData.at(41), "o_modified_adress");
	updataValue(tempSql, rowData.at(42), "o_abnormal_info");
	updataValue(tempSql, rowData.at(43), "o_tmall_voucher");
	updataValue(tempSql, rowData.at(44), "o_jifenbao_voucher");
	updataValue(tempSql, rowData.at(45), "o_o2o_trading");
	updataValue(tempSql, rowData.at(46), "o_trading_type");
	updataValue(tempSql, rowData.at(47), "o_retailshop_name");
	updataValue(tempSql, rowData.at(48), "o_retailshop_id");
	updataValue(tempSql, rowData.at(49), "o_retaildelivery_name");
	updataValue(tempSql, rowData.at(50), "o_retaildelivery_id");

	updataValue(tempSql, rowData.at(51), "o_refund_account");
	updataValue(tempSql, rowData.at(52), "o_appointment_shop");
	updataValue(tempSql, rowData.at(53), "b_order_confirmtime");
	updataValue(tempSql, rowData.at(54), "b_pay_confirmaccount");
	updataValue(tempSql, rowData.at(55), "o_buyer_envelope");
	updataValue(tempSql, rowData.at(56), "o_mainorder_indexcode");
	updataValue(tempSql, rowData.at(57), "o_ext1_info");
	updataValue(tempSql, rowData.at(58), "o_ext2_info");

	tempSql += ", `updatatime` = NOW()";
	tempSql += ", `versions` = `versions`+1";

	sql = tempSql + tempValues;
}

bool OrderManagement::isExistOrder(const QString& indexCode)
{
	QSqlQuery query(*m_mySqlDB);
	QString sql = QString("SELECT * FROM `order` WHERE order_indexcode = %1").arg(indexCode);
	query.exec(sql);
	return query.next();
}

void OrderManagement::addValue(QString& sql, QString& values, const QVariant& var, const QString& marke)
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

void OrderManagement::updataValue(QString& sql, const QVariant& var, const QString& marke)
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

void OrderManagement::readExcelData(const QString& excelFilePath, ExcelList& excelList)
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
	castVariant2ListListVariant(var, excelList);
}

void OrderManagement::openExcelFile()
{
	QString xlsFile = QFileDialog::getOpenFileName(this, QString(), QString(), "excel(*.xls *.xlsx)");
	if (xlsFile.isEmpty())
		return;
	qDebug() << "open file :" << xlsFile;
	ExcelList excelList;
	readExcelData(xlsFile, excelList);
	//todo
	if(excelList.at(0).size() == 11)
		updataCommodity(excelList);
	if (excelList.at(0).size() > 50)
		updataOrder(excelList);

}

void OrderManagement::openDBInfoWidget()
{
	if (m_mySqlInfo)
		m_mySqlInfo->show();
}

void OrderManagement::on_openMySql(const QString &host, const QString &user, const QString &password)
{
	closeMySqlDB();
	openMySqlDB(host, "web_shop", user, password);
}
