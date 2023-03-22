#include <QtGui>
#include <QAxObject>
#include <QApplication>
#include <QStandardPaths>
#include <QFileDialog>
#include "configtab.h"

#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    ConfigTab w;
    w.show();

    return a.exec();
}
