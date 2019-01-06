#include "OrderManagement.h"
#include <QtWidgets/QApplication>

int main(int argc, char *argv[])
{
	QApplication a(argc, argv);
	OrderManagement w;
	w.show();
	return a.exec();
}
