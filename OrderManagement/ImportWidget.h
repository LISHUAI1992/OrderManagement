#pragma once

#include <QWidget>
#include "ui_ImportWidget.h"

class ImportWidget : public QWidget
{
	Q_OBJECT

public:
	ImportWidget(QWidget *parent = Q_NULLPTR);
	~ImportWidget();

	void SetBarValue(int value);

private slots:
	void  on_ExeSqlResult(const QString &result);

private:
	Ui::ImportWidget ui;
	int m_value;
};
