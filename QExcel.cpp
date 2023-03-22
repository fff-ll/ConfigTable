#include "qexcel.h"
#include <QDir>
#include <QDateTime>

QExcel::QExcel(QObject *parent):QObject(parent)
{
    excel = nullptr;
    workBooks = nullptr;
    workBook = nullptr;
    sheets = nullptr;
    sheet = nullptr;
}

QExcel::~QExcel()
{

}

/***************************************************************************/
/* 文件操作                                                                 */
/**************************************************************************/

void QExcel::open(QString FileName)
{
    if(excel==nullptr)
    {
        excel = new QAxObject();
        excel->setControl("Excel.Application");                            //连接Excel控件
        excel->setProperty("DisplayAlerts", false);                        //禁止显示警告
        workBooks = excel->querySubObject("Workbooks");                    //获取工作簿集合
        QFile file(FileName);
        if(file.exists())
        {
            workBooks->dynamicCall("Open(const QString&)", FileName);      //打开指定文件
            workBook = excel->querySubObject("ActiveWorkBook");            //获取活跃工作簿
            sheets = workBook->querySubObject("WorkSheets");               //获取工作表集合
            sheet = workBook->querySubObject("ActiveSheet");               //获取活跃工作表
        }
    }
    else
    {
        workBooks = excel->querySubObject("Workbooks");                    //获取工作簿集合
        QFile file(FileName);
        if(file.exists())
        {
            workBooks->dynamicCall("Open(const QString&)", FileName);      //打开指定文件
            workBook = excel->querySubObject("ActiveWorkBook");            //获取活跃工作簿
            sheets = workBook->querySubObject("WorkSheets");               //获取工作表集合
            sheet = workBook->querySubObject("ActiveSheet");               //获取活跃工作表
        }
    }


//    QString merge_cell;
//    merge_cell.append(QChar('A'));  //初始列
//    merge_cell.append(QString::number(1));  //初始行
//    merge_cell.append(":");
//    merge_cell.append(QChar(getUsedColumnCount() - 1 + 'A'));  //终止列
//    merge_cell.append(QString::number(getUsedRowsCount() + 1));  //终止行
//    QAxObject *merge_range = sheet->querySubObject("Range(const QString&)", merge_cell);
//    merge_range->setProperty("NumberFormat", "@");
}

void QExcel::createFile(QString FileName)
{
    if(excel==nullptr)
    {
        excel = new QAxObject();
        excel->setControl("Excel.Application");                            //连接Excel控件
        excel->setProperty("DisplayAlerts", false);                        //禁止显示警告
        workBooks = excel->querySubObject("Workbooks");                    //获取工作簿集合
        QFile file(FileName);
        if(!file.exists())
        {
            workBooks->dynamicCall("Add");                                 //新建文件
            workBook=excel->querySubObject("ActiveWorkBook");              //获取活跃工作簿
            workBook->dynamicCall("SaveAs(const QString&)",FileName);      //按指定文件名保存文件
            sheets = workBook->querySubObject("WorkSheets");               //获取工作表集合
            sheet = workBook->querySubObject("ActiveSheet");               //获取活跃工作表
        }
    }
    else
    {
        workBooks = excel->querySubObject("Workbooks");                    //获取工作簿集合
        QFile file(FileName);
        if(!file.exists())
        {
            workBooks->dynamicCall("Add");                                 //新建文件
            workBook=excel->querySubObject("ActiveWorkBook");              //获取活跃工作簿
            workBook->dynamicCall("SaveAs(const QString&)",FileName);      //按指定文件名保存文件
            sheets = workBook->querySubObject("WorkSheets");               //获取工作表集合
            sheet = workBook->querySubObject("ActiveSheet");               //获取活跃工作表
        }
    }
}

void QExcel::save() //保存Excel文件
{
//    qDebug("0%s",qPrintable(filename));
//    workBook->dynamicCall("SaveAs(const QString &)", filename);;
     workBook->dynamicCall("Save()");
    qDebug()<<"文件已保存";
}

void QExcel::quit() //退出Excel
{
    excel->dynamicCall("Quit()");

    delete sheet;
    delete sheets;
    delete workBook;
    delete workBooks;
    delete excel;

    excel = nullptr;
    workBooks = nullptr;
    workBook = nullptr;
    sheets = nullptr;
    sheet = nullptr;

    qDebug()<<"退出Excel";
}

/**************************************************************************/
/* 工作表操作                                                              */
/*************************************************************************/

void QExcel::selectSheet(const QString& sheetName) //根据名字选择工作表
{
    sheet = sheets->querySubObject("Item(const QString&)", sheetName);
}

void QExcel::selectSheet(int sheetIndex) //根据下标索引选择工作表，下标从1开始
{
    sheet = sheets->querySubObject("Item(int)", sheetIndex);
}

void QExcel::deleteSheet(const QString& sheetName) //根据名字删除工作表
{
    QAxObject * a = sheets->querySubObject("Item(const QString&)", sheetName);
    a->dynamicCall("delete");
}

void QExcel::deleteSheet(int sheetIndex) //根据下标索引删除工作表，下标从1开始
{
    QAxObject * a = sheets->querySubObject("Item(int)", sheetIndex);
    a->dynamicCall("delete");
}

void QExcel::insertSheet(QString sheetName) //新建工作表
{
    sheets->querySubObject("Add()");
    QAxObject * a = sheets->querySubObject("Item(int)", 1);
    a->setProperty("Name", sheetName);
}

int QExcel::getSheetsCount() //获取工作表数量
{
    return sheets->property("Count").toInt();
}

QString QExcel::getSheetName() //获取当前活跃工作表的名字
{
    return sheet->property("Name").toString();
}

QString QExcel::getSheetName(int sheetIndex) //根据下标索引获取工作表名字
{
    QAxObject * a = sheets->querySubObject("Item(int)", sheetIndex);
    return a->property("Name").toString();
}

/***************************************************************************/
/* 单元格操作                                                               */
/**************************************************************************/

//void QExcel::mergeCells(int topLeftRow, int topLeftColumn, int bottomRightRow, int bottomRightColumn) //根据行列编号合并单元格
//{
//    QString cell;
//    cell.append(QChar(topLeftColumn - 1 + 'A'));
//    cell.append(QString::number(topLeftRow));
//    cell.append(":");
//    cell.append(QChar(bottomRightColumn - 1 + 'A'));
//    cell.append(QString::number(bottomRightRow));

//    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
//    range->setProperty("VerticalAlignment", -4108);//xlCenter
//    range->setProperty("WrapText", true);
//    range->setProperty("MergeCells", true);
//}

//void QExcel::mergeCells(const QString& cell) //根据字母编号合并单元格
//{
//    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
//    range->setProperty("VerticalAlignment", -4108);//xlCenter
//    range->setProperty("WrapText", true);
//    range->setProperty("MergeCells", true);
//}

void QExcel::setCellString(int row, int column,QVariant value) //根据行列编号设置单元格数据
{
    QAxObject *range = sheet->querySubObject("Cells(int,int)", row, column);
    range->dynamicCall("SetValue(QVariant)", value);
    workBook->dynamicCall("Save()");
}

void QExcel::setCellString(const QString& cell, const QString& value) //根据字母编号设置单元格数据
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->dynamicCall("SetValue(const QString&)", value);
}

bool QExcel::setRowData(int row,int colum_start,int column_end,QVariant vData) //批量写入一行数据
{
    bool op = false;
    QString start,end;
    start=getRange_fromColumn(colum_start);
    end=getRange_fromColumn(column_end);
    QVariant qstrRange = start+QString::number(row,10)+":"+end+QString::number(row,10);
    QAxObject *range = sheet->querySubObject("Range(const QString&)", qstrRange);
    if ( range )
    {
        range->dynamicCall("SetValue(const QVariant&)",QVariant(vData)); //修改单元格的数据
        op = true;
    }
    else
    {
        op = false;
    }

    delete range;
    return op;
}

QString QExcel::getRange_fromColumn(int column) //根据列值计算出字母值
{
    if(column <= 0) //列值必须大于等于1
        return "";

    QString ABC="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    QString result="";
    QVector<int> pos;
    if(column>=1 && column<=26)
    {
        result += ABC[column-1];
        return result;
    }
    else
    {
        int high = column;
        int low;
        int last_high;
        while(high>0)
        {
            last_high=high;
            high = high / 26;
            low = last_high % 26;
            if(low==0 && high!=0)
            {
                low=26;
                high=high-1;
            }
            pos.push_front(low);
        }

        for(int i=0; i<pos.size(); i++)
        {
            result += ABC[pos[i]-1];
        }
    }
    return result;
}

QVariant QExcel::getCellValue(int row, int column) //根据行列编号获取单元格数值
{
    QAxObject *range = sheet->querySubObject("Cells(int,int)", row, column);

    bool bMerger = range->property("MergeCells").toBool();
    if (bMerger)
    {
        range = range->querySubObject("MergeArea");
        int nRowStart = range->property("Row").toInt();//左上角x
        int nRowEnd = range->property("Column").toInt();  //左上角y
        range =  sheet->querySubObject("Cells(int,int)", nRowStart, nRowEnd);
    }

    return range->property("Value");
}

QVariant QExcel::getCellValue(const QString& cell) //根据字母编号获取单元格数值
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);


    bool bMerger = range->property("MergeCells").toBool();
    if (bMerger)
    {
        range = range->querySubObject("MergeArea");
        int nRowStart = range->property("Row").toInt();//左上角x
        int nRowEnd = range->property("Column").toInt();  //左上角y
        range =  sheet->querySubObject("Cells(int,int)", nRowStart, nRowEnd);
    }

    return range->property("Value");
}

void QExcel::clearCell(int row, int column) //根据行列编号清空单元格
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->dynamicCall("ClearContents()");
}

void QExcel::clearCell(const QString& cell) //根据字母编号清空单元格
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->dynamicCall("ClearContents()");
}

void QExcel::getUsedRange(int *topLeftRow, int *topLeftColumn, int *bottomRightRow, int *bottomRightColumn) //取得工作表已使用范围
{
    QAxObject *usedRange = sheet->querySubObject("UsedRange");
    *topLeftRow = usedRange->property("Row").toInt();
    *topLeftColumn = usedRange->property("Column").toInt();

    QAxObject *rows = usedRange->querySubObject("Rows");
    *bottomRightRow = *topLeftRow + rows->property("Count").toInt() - 1;

    QAxObject *columns = usedRange->querySubObject("Columns");
    *bottomRightColumn = *topLeftColumn + columns->property("Count").toInt() - 1;
}

int QExcel::getUsedRowsCount() //获取总行数
{
    QAxObject *usedRange = sheet->querySubObject("UsedRange");
    int topRow = usedRange->property("Row").toInt();
    QAxObject *rows = usedRange->querySubObject("Rows");
    int bottomRow = topRow + rows->property("Count").toInt() - 1;
    return bottomRow;
}

int QExcel::getUsedColumnCount() //获取总列数
{
    QAxObject *usedRange = sheet->querySubObject("UsedRange");
    int leftColumn = usedRange->property("Column").toInt();
    QAxObject *columns = usedRange->querySubObject("Columns");
    int rightColumn = leftColumn + columns->property("Count").toInt() - 1;
    return rightColumn;
}

QStringList QExcel::getconfigexcelname(QString value,QString eventtype)//获取要更新的配置表名称
{
    if(eventtype=="排位赛" || eventtype=="bp" || eventtype=="小活动" || eventtype=="矿活动"){
        eventtype="主线活动";
    }
    selectSheet(eventtype);
    QStringList configexcelname;
    int row=getconfig(value);
    QAxObject *range = sheet->querySubObject("Cells(int,int)", row, 1);
    bool bMerger = range->property("MergeCells").toBool();
    if (bMerger)
    {
          range = range->querySubObject("MergeArea");
          QVariant var = range->dynamicCall("Value");
          int rows=var.toList().size();
          for(int i=0;i<rows;i++){
              configexcelname.append(getCellValue(row+i,2).toString());
          }
    }else{
        configexcelname.append(getCellValue(row,2).toString());
    }
    return configexcelname;
}

int QExcel::getconfig(QString value) //获取活动名所在行
{
    int row=getUsedRowsCount();
   for(int i=1;i<=row;i++){
       if(value==getCellValue(i,1)){
           return i;
       };
   }
   return -1;
}

QString QExcel::getstarttime(int row){
    QString maxtime="2022-00-00 00:00:00";
    for(int i=2;i<=row;i++){
        QString cur=getCellValue(i,9).toString();
        if(cur>=maxtime){
            maxtime=cur;
        }
    }
    maxtime[10]=' ';
    QDateTime dateTime= QDateTime::currentDateTime();//获取系统当前的时间
    QString str = dateTime .toString("yyyy-MM-dd hh:mm:ss");
    if(maxtime=="2022-00-00 00:00:00" || str>maxtime){
        return str;
    }
    QDateTime time_max=QDateTime::fromString(maxtime.mid(0,19), "yyyy-MM-dd hh:mm:ss");
    maxtime= time_max.addDays(1).toString("yyyy-MM-dd hh:mm:ss");
    return maxtime;
}

QString QExcel::add_time(QString starttime,int days){
    QDateTime time_end=QDateTime::fromString(starttime, "yyyy-MM-dd hh:mm:ss");
    QString endtime= time_end.addDays(days).toString("yyyy-MM-dd hh:mm:ss");
    return endtime;
}

void QExcel::seteventconfig(QString eventtype,QString ename,QString econfig,QString starttime){
    int row=getUsedRowsCount()+1;//最下面一行
    int id=getCellValue(row-1,1).toInt();
    QString sid=QString::number(id+1,10);
    setCellString(row,1,sid);//添加id
    setCellString(row,2,ename);
    setCellString(row,3,econfig);//添加两项配置

    if(eventtype=="副本"){
        setCellString(row,7,"-12:00:00");
    }//添加副本预告

    setCellString(row,8,starttime);
    QAxObject *cell =sheet->querySubObject("Cells(int,int)", row, 8);
    cell->setProperty("NumberFormatLocal", "yyyy-mm-dd hh:mm:ss");
    setCellString(row,9,add_time(starttime,4));
    QAxObject *cellend =sheet->querySubObject("Cells(int,int)", row, 9);
    cellend->setProperty("NumberFormatLocal", "yyyy-mm-dd hh:mm:ss");//添加开始结束时间

    if(eventtype=="排位赛"){
        QDateTime time_r=QDateTime::fromString(add_time(starttime,5), "yyyy-MM-dd hh:mm:ss");
        QString rtime= time_r.addSecs(-7200).toString("yyyy-MM-dd hh:mm:ss");
        setCellString(row,11,rtime);
        QAxObject *cellr =sheet->querySubObject("Cells(int,int)", row, 11);
        cellr->setProperty("NumberFormatLocal", "yyyy-mm-dd hh:mm:ss");
    }//添加排位赛释放时间

    if(eventtype=="主线活动"){
        setCellString(row,13,"14");
    }else{
        setCellString(row,13,"7");
    }//添加等级
}

void QExcel::copycolumn_sheet(QString sheet1, QString start1, QString sheet2, QString start2, int rows, int cols)//工作表之间复制一定区域
{
    selectSheet(sheet1);
    QAxObject* Range_1 = sheet->querySubObject ("Range(QString)",start1);
    Range_1 = Range_1->querySubObject("Resize(int,int)",rows,cols);
    QVariant var_1 = Range_1->dynamicCall("Value");

    selectSheet(sheet2);
    Range_1 = sheet->querySubObject ("Range(QString)",start2);
    Range_1 = Range_1->querySubObject("Resize(int,int)",rows,cols);
    Range_1->setProperty ("Value",var_1);

    workBook->dynamicCall("Save()");
}

void QExcel::getconfigname(){
    selectSheet("副本");
    int row_1=getUsedRowsCount();
    copycolumn_sheet("副本","A1","活动名称","A2",row_1,1);

    copycolumn_sheet("主线活动","A1","活动名称","G2",5,1);//小活动
    copycolumn_sheet("主线活动","A6","活动名称","D2",5,1);//矿
    copycolumn_sheet("主线活动","A11","活动名称","E2",5,1);//排位赛
    copycolumn_sheet("主线活动","A16","活动名称","F2",5,1);//bp

    selectSheet("主线活动");
    int row_6=getUsedRowsCount()-20;
    copycolumn_sheet("主线活动","A21","活动名称","B2",row_6,1);//主线活动

    selectSheet("其余功能");
    int row_3=getUsedRowsCount();
    copycolumn_sheet("其余功能","A1","活动名称","C2",row_3,1);//其余功能

}

QString QExcel::geteventType(QString eventname)//获取活动类型
{
    selectSheet("活动名称");
    QAxObject *usedRange = sheet->querySubObject("UsedRange");
    QVariant var = usedRange->dynamicCall("Value");
    QList< QList<QVariant> >  datas;//构建数据模型用于存储数据
    Qvariant2listlistVariant(var,datas);
    for (int a=0;a<datas.size();a++ )
    {
       for (int b=0;b<datas.at(a).size();b++ )
       {
           if(eventname==QString(datas[a][b].toByteArray())){
               return getCellValue(1,b+1).toString();
       }
       }
    }
    return "";

}

//将Qvariant转换为QList<QList<QVariant>>
void QExcel::Qvariant2listlistVariant(const QVariant &var,QList<QList<QVariant>> &ret)
{
    QVariantList varrows=var.toList();
    if(varrows.isEmpty())
    {
        return;
    }
    else {
        const int rowcount=varrows.size();//行数
        QVariantList rowdata;
        for(int i=0;i<rowcount;i++)
        {
            rowdata=varrows[i].toList();//将每一行的值存入到list中
            ret.push_back(rowdata);
        }
    }
}
QAxObject *QExcel::getWorkBooks()
{
    return workBooks;
}

QAxObject *QExcel::getWorkBook()
{
    return workBook;
}

QAxObject *QExcel::getWorkSheets()
{
    return sheets;
}

QAxObject *QExcel::getWorkSheet()
{
    return sheet;
}



