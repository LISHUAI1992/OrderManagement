#include "OrderManagement.h"

OrderManagement::OrderManagement(QWidget *parent)
	: QMainWindow(parent)
	, m_mySqlInfo(NULL)
	, m_pCore(NULL)
{
	ui.setupUi(this);
	m_pCore = new OrderCore();
	setWindowTitle(TU("订单管理系统"));
	m_mySqlInfo = new MySQLInfo();
	m_mySqlInfo->hide();

	ui.actionupdatafile->setText(TU("上传文件"));
	ui.actionsetDB->setText(TU("设置数据库信息"));
	
	connect(ui.actionupdatafile, SIGNAL(triggered()),
		this, SLOT(openExcelFile()));
	connect(ui.actionsetDB, SIGNAL(triggered()),
		this, SLOT(openDBInfoWidget()));
	connect(m_mySqlInfo, SIGNAL(openMysql(const QString &, const QString &, const QString &)),
		this, SLOT(on_openMySql(const QString &, const QString &, const QString &)));
	//m_pCore->openMySqlDB("132.232.101.227", "shop", "myuser", "Hik19920623#123");
	m_pCore->OpenMySqlDB("127.0.0.1", "web_shop", "root", "12345");
}

OrderManagement::~OrderManagement()
{
	SAFEDELETE(m_mySqlInfo);
	SAFEDELETE(m_pCore);
}

void OrderManagement::openExcelFile()
{
	QString xlsFile = QFileDialog::getOpenFileName(this, QString(), QString(), "excel(*.xls *.xlsx)");
	if (xlsFile.isEmpty())
		return;
	qDebug() << "open file :" << xlsFile;
	ExcelList excelList;
	m_pCore->ReadExcelData(xlsFile, excelList);
	//todo
	if(excelList.at(0).size() == 11)
		m_pCore->UpdataCommodity(excelList);
	if (excelList.at(0).size() > 50)
		m_pCore->UpdataOrder(excelList);

}

void OrderManagement::openDBInfoWidget()
{
	if (m_mySqlInfo)
		m_mySqlInfo->show();
}

void OrderManagement::on_openMySql(const QString &host, const QString &user, const QString &password)
{
	m_pCore->CloseMySqlDB();
	m_pCore->OpenMySqlDB(host, "web_shop", user, password);
}
