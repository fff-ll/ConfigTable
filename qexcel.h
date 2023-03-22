#ifndef QEXCEL_H
#define QEXCEL_H

#include <QString>
#include <QColor>
#include <QVariant>
#include <QAxObject>
#include <QFile>
#include <QStringList>
#include <QDebug>

class QAxObject;

class QExcel : public QObject
{
public:
    QExcel(QObject *parent = nullptr);
    ~QExcel();

public:
    QAxObject * getWorkBooks();
    QAxObject * getWorkBook();
    QAxObject * getWorkSheets();
    QAxObject * getWorkSheet();
    QAxObject * excel;         //Excel指针
    QAxObject * workBooks;     //工作簿集合
    QAxObject * workBook;      //工作簿
    QAxObject * sheets;        //工作表集合
    QAxObject * sheet;         //工作表

public:
    /***************************************************************************/
    /* 文件操作                                                                 */
    /**************************************************************************/

    void open(QString FileName);                   //打开文件
    void createFile(QString FileName);             //创建文件
    void save();                                   //保存Excel文件
    void quit();                                  //退出Excel

    /**************************************************************************/
    /* 工作表操作                                                              */
    /*************************************************************************/
    void selectSheet(const QString& sheetName);   //根据名字选择工作表
    void selectSheet(int sheetIndex);             //根据下标索引选择工作表，下标从1开始
    void deleteSheet(const QString& sheetName);   //根据名字删除工作表
    void deleteSheet(int sheetIndex);             //根据下标索引删除工作表，下标从1开始
    void insertSheet(QString sheetName);          //新建工作表
    int getSheetsCount();                         //获取工作表数量
    QString getSheetName();                       //获取当前活跃工作表的名字
    QString getSheetName(int sheetIndex);         //根据下标索引获取工作表名字
    void copycolumn_sheet(QString sheet1,QString start1,QString sheet2,QString start2,int rows,int cols);//工作表之间复制一定区域


    /***************************************************************************/
    /* 单元格操作                                                               */
    /**************************************************************************/
    QStringList getconfigexcelname(QString value,QString eventtype);//获取要更新的配置表名称
    int getconfig(QString value);//获取配置所在行
    void seteventconfig(QString eventtype,QString ename,QString econfig,QString starttime);//添加副本配置
    QString getstarttime(int row);
    QString add_time(QString starttime,int days);//获取最大时间
    void getconfigname();//生成活动名字表格
    QString geteventType(QString eventname);//获取活动类型
    void setCellString(int row, int column, QVariant value); //根据行列编号设置单元格数据，例如(1,1,"xxx")
    void setCellString(const QString& cell, const QString& value); //根据字母编号设置单元格数据，例如("A5","xxx")
    bool setRowData(int row,int colum_start,int column_end,QVariant vData);  //批量写入一行数据
    QString getRange_fromColumn(int column);                                 //根据列值计算出字母值
    QVariant getCellValue(int row, int column);                    //根据行列编号获取单元格数值
    QVariant getCellValue(const QString& cell);                    //根据字母编号获取单元格数值
    void clearCell(int row, int column);                           //根据行列编号清空单元格
    void clearCell(const QString& cell);                           //根据字母编号清空单元格
    void getUsedRange(int *topLeftRow, int *topLeftColumn, int *bottomRightRow, int *bottomRightColumn); //取得工作表已使用范围
    int getUsedRowsCount();                                        //获取总行数
    int getUsedColumnCount();                                      //获取总列数

    //将Qvariant转换为QList<QList<QVariant>>
    void Qvariant2listlistVariant(const QVariant &var,QList<QList<QVariant>> &ret);

};

#endif


