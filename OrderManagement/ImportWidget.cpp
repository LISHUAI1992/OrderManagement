#include "ImportWidget.h"

ImportWidget::ImportWidget(QWidget *parent)
	: QWidget(parent)
	, m_value(0)
{
	ui.setupUi(this);
}

ImportWidget::~ImportWidget()
{
}

void ImportWidget::SetBarValue(int value)
{
	ui.progressBar->setRange(0, value);
	m_value = 0;
	ui.textEdit->clear();
}

void ImportWidget::on_ExeSqlResult(const QString &result)
{
	m_value += 1;
	ui.progressBar->setValue(m_value);
	ui.textEdit->append(result);
}
