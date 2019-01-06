#include "MySQLInfo.h"

MySQLInfo::MySQLInfo(QWidget *parent)
	: QWidget(parent)
	, m_hostName("132.232.101.227")
	, m_passWord("Hik19920623#123")
	, m_userName("myuser")
{
	setWindowFlags(windowFlags() &~(Qt::WindowMinMaxButtonsHint));
	setWindowTitle(TU("Êý¾Ý¿âÉèÖÃ"));
	ui.setupUi(this);
	ui.hostEdit->setText(m_hostName);
	ui.userEdit->setText(m_userName);
	ui.passwordEdit->setEchoMode(QLineEdit::Password);
	ui.passwordEdit->setText(m_passWord);

}

MySQLInfo::~MySQLInfo()
{
}

void MySQLInfo::on_saveBtn_clicked()
{
	m_hostName = ui.hostEdit->text();
	m_userName = ui.userEdit->text();
	m_passWord = ui.passwordEdit->text();

	emit openMysql(m_hostName, m_userName, m_passWord);
}
