#ifndef CONFIGTAB_H
#define CONFIGTAB_H

#include <QMainWindow>
#include <QWidget>
#include <QAxObject>
#include <QDebug>
#include <QTableWidget>
#include <QCheckBox>
#include "qexcel.h"
#include <QProcess>
#include <QStackedLayout>

QT_BEGIN_NAMESPACE
namespace Ui { class ConfigTab; }
QT_END_NAMESPACE

class ConfigTab : public QMainWindow
{
    Q_OBJECT
private:
    Ui::ConfigTab *ui;
    QString eventname;
    QList<QCheckBox*> checkboxlist;//动态创建按钮的列表
    QString path;
    QStackedLayout* slayout;
//    int m_totalRowCnt;
public:
    ConfigTab(QWidget *parent = nullptr);
    ~ConfigTab();
    QExcel * event;
    void tab_update();
    void excel_read(QTableWidget *tablewidget);
    void TableWidgetAddCheckBox(QTableWidget *tableWidget, int x, int y, QString text, Qt::CheckState checkState);
    QList<QString> addconfig(QList<QString> configlist);//批量添加配置
    bool copycolumn_excel(QString sheet,int row,int col,QString excelname,int rows,int cols);//表格之间复制一定区域
    void catend(QList<QString> result);
private slots:
    void on_action11_triggered();
    void sloton_checkBox_stateChanged(int arg1);//动态创建信号槽函数
    void on_add_config_clicked();
    void on_action22_triggered();
    void processFinished(int, QProcess::ExitStatus status);

//signals:
//    void lack();
//    void complete();
};
#endif // CONFIGTAB_H
