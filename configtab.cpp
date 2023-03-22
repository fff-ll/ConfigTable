#include "configtab.h"
#include "ui_ConfigTab.h"
#include "qexcel.h"

#include <QFileDialog>
#include <QVariant>
#include <QFile>
#include <QDir>
#include <QTableWidget>
#include <QCheckBox>
#include <QHBoxLayout>
#include <QTextBlock>
#include <QVector>

#include <QFile>
#include <QUrl>

#include <QProcess>
#include <windows.h>
#include <QMovie>
#include <QLabel>
#include <QStackedLayout>
#include <QVBoxLayout>
#include <QThread>
#include <QtConcurrent>
#include <QTime>



ConfigTab::ConfigTab(QWidget *parent) : QMainWindow(parent) , ui(new Ui::ConfigTab)
{
    ui->setupUi(this);
    ui->textEdit->setReadOnly(true);
    ui->textEdit_2->setReadOnly(true);
    ui->tab_config->setEditTriggers(QAbstractItemView::NoEditTriggers);
//配置表路径
    QString str = QApplication::applicationDirPath();
    path=str+"/config";
//创建tablewidget
    event=new QExcel();
    tab_update();
//设置布局
    ui->centralwidget->setLayout(ui->verticalLayout_3);
    ui->verticalLayout_3->setMargin(5);
//创建loading遮罩
    ui->shadow->setFixedSize(size());
    QPalette pal(ui->shadow->palette());
    pal.setColor(QPalette::Background, QColor(0, 0, 0, 50));
    ui->shadow->setAutoFillBackground(true);
    ui->shadow->setPalette(pal);
    ui->shadow->stackUnder(this);
    ui->shadow->setAlignment(Qt::AlignCenter);
    QMovie *loading=new QMovie(str+"/loading.gif");
    ui->shadow->setMovie(loading);
    ui->shadow->raise();
    ui->shadow->hide();
//创建添加配置等待动画
    QMovie *addloading=new QMovie(str+"/addloading.gif");
    addloading->start();
    ui->label->setMovie(addloading);
    ui->label->setVisible(false);

    qDebug("创建完成");

}

ConfigTab::~ConfigTab()
{

    event->quit();
    delete event;
    delete ui;
}


void ConfigTab::on_action11_triggered()
{
    ui->shadow->movie()->start();
    ui->shadow->show();

    event->quit();
    QProcess *process=new QProcess();
    process->setCreateProcessArgumentsModifier([](QProcess::CreateProcessArguments *args)
    {
        args->startupInfo->wShowWindow = SW_HIDE;
        args->startupInfo->dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
    }
    );
    connect(process, SIGNAL(finished(int,QProcess::ExitStatus)), this, SLOT(processFinished(int, QProcess::ExitStatus)));
    QString str = QApplication::applicationDirPath();
    str += "/down/down.exe";
    process->start(str);

}

void ConfigTab::processFinished(int, QProcess::ExitStatus status)
{
    if(status==QProcess::NormalExit){
        tab_update();
        ui->shadow->hide();
        ui->textEdit_2->append("更新成功");
     }else{
        ui->shadow->hide();
        ui->textEdit_2->append("下载失败，请重试");
        return;
    }

}

void ConfigTab::on_action22_triggered()
{
    event->quit();
    tab_update();
    ui->textEdit_2->append("更新成功");
}


void ConfigTab::tab_update(){

    event->open(path+"/配置汇总.xlsx");
    event->getconfigname();
    event->selectSheet("活动名称");

    ui->tab_config->setRowCount(event->getUsedRowsCount()-1);
    ui->tab_config->setColumnCount(event->getUsedColumnCount());
    QStringList m_Header;
    m_Header<<QString("副本")<<QString("主线活动")<<QString("其余功能")<<QString("矿活动")<<QString("排位赛")<<QString("bp")<<QString("小活动");
    ui->tab_config->setHorizontalHeaderLabels(m_Header);//添加横向表头
    ui->tab_config->verticalHeader()->setVisible(true);//纵向表头可视化
    ui->tab_config->horizontalHeader()->setVisible(true);//横向表头可视化

    excel_read(ui->tab_config);
}

void ConfigTab::excel_read(QTableWidget *tablewidget)
{
    checkboxlist.clear();
    QAxObject *usedrange=event->sheet->querySubObject("UsedRange");//获取整个worksheet的范围
    QVariant cell=usedrange->dynamicCall("value");//将worksheet的全部内容存储在QVariant变量中
    QList<QList<QVariant>> ret;//将Qvariant转换为QList<QList<QVariant>>
    event->Qvariant2listlistVariant(cell,ret);

    for(int j=0;j<tablewidget->columnCount();j++)
    {
        int cur=0;
        for(int i=0;i<tablewidget->rowCount();i++)
        {
            QString strVal=ret.at(i+1).at(j).toString();
            //将单元格的内容放置在table表中
//            tablewidget->setItem(i,j,new QTableWidgetItem(strVal));

            if(strVal!=nullptr){
                TableWidgetAddCheckBox(tablewidget,cur,j,strVal,Qt::Unchecked);
                cur++;
            }
//            tablewidget->item(i,j)->setTextAlignment(Qt::AlignVCenter|Qt::AlignHCenter);
        }
    }

}

//TableWidget中加入复选框
void ConfigTab::TableWidgetAddCheckBox(QTableWidget *tableWidget, int x, int y, QString text, Qt::CheckState checkState)
{
    QWidget *widget = new QWidget();
    QHBoxLayout *layout = new QHBoxLayout();
    QCheckBox *checkBox = new QCheckBox;

    checkBox->setText(text);                		//复选框文本
    checkBox->setCheckState(checkState);   			//复选框初始状态
    layout->addWidget(checkBox,0,Qt::AlignCenter);	//居中
    layout->setMargin(0);							//左右间距
    widget->setLayout(layout);
    widget->setLayoutDirection(Qt::RightToLeft);
    tableWidget->setCellWidget(x,y,widget);
    //向QList添加成员项
    checkboxlist.append(checkBox);
    //关联信号槽函数
    connect(checkboxlist.last(),SIGNAL(stateChanged(int)),this,SLOT(sloton_checkBox_stateChanged(int)));
}


void ConfigTab::sloton_checkBox_stateChanged(int arg1)
{
    QCheckBox *cb = qobject_cast<QCheckBox*>(sender());

    if(arg1==Qt::Checked){
        QString cur_text=cb->text();
        ui->textEdit->append(cur_text);
    }
    if(arg1==Qt::Unchecked){
        QString cur_text=cb->text();
        ui->textEdit->moveCursor(QTextCursor::Start);//移动到开头
        ui->textEdit->find(cur_text);//寻找

        ui->textEdit->textCursor().removeSelectedText();//删除
        ui->textEdit->moveCursor(QTextCursor::NextBlock);//移动到下行开始
        ui->textEdit->textCursor().deletePreviousChar();//删除空行
    }
}

void ConfigTab::on_add_config_clicked()
{

    int num=ui->textEdit->document()->lineCount();
    ui->label->setVisible(true);

    QList<QString> configlist;
    ui->textEdit_2->append("添加");
    for (int i=0;i<num;i++) {
        QString cur=ui->textEdit->document()->findBlockByLineNumber(i).text();
        ui->textEdit_2->insertPlainText(cur+"，");
        configlist.append(cur);
    }
    QFuture<QList<QString>> fu=QtConcurrent::run(this,&ConfigTab::addconfig,configlist);

    QElapsedTimer t;
    t.start();
    while (!fu.isFinished()) {
        QApplication::processEvents();
    }

    QList<QString> result=fu.result();
    catend(result);

    ui->textEdit->clear();
    for (int i=0;i<checkboxlist.size();i++) {
        if(checkboxlist[i]->checkState()==Qt::Checked){
            checkboxlist[i]->setCheckState(Qt::Unchecked);
        }
    }

}

void ConfigTab::catend(QList<QString> result){

    if(result.size()==0){
        ui->textEdit_2->insertPlainText("已添加");
        ui->label->setVisible(false);
        return;
    }
    ui->textEdit_2->insertPlainText("更新配置表");
    if(result.last()=="0"){
        ui->textEdit_2->append("...添加失败，缺少"+result[0]);
    }else{
        for (int i=0;i<result.size();i++) {
            ui->textEdit_2->insertPlainText("\""+result[i]+"\"");
        }
    }
    ui->label->setVisible(false);
}

QList<QString> ConfigTab::addconfig(QList<QString> configlist){
    CoInitializeEx(NULL, COINIT_MULTITHREADED);

    QExcel *event1=new QExcel();
    event1->open(path+"/配置汇总.xlsx");

    QList<QString> log_excelname;

    QStringList eventtypelist;
    eventtypelist<<QString("副本")<<QString("主线活动")<<QString("其余功能")<<QString("矿活动")<<QString("排位赛")<<QString("bp")<<QString("小活动");
    QVector<int> eventtype_timelist(7,0);//对每种类型活动的时间分别处理

    QExcel* all_event_time=new QExcel();
    all_event_time->open(path+"/all_event_time.xlsx");//打开配置表
    QString starttime=all_event_time->getstarttime(all_event_time->getUsedRowsCount());//获得本次添加活动前的最大时间

    for (int i=0;i<configlist.size();i++) { //循环处理活动
        QString eventname=configlist[i];
        QString eventtype=event1->geteventType(eventname);

        QStringList configexcelname=event1->getconfigexcelname(eventname,eventtype);//获得活动需要更新的excel表名列表
        int row=event1->getconfig(eventname);
        for (int i=0;i<configexcelname.size();i++) { //循环更新excel表

            QFile file(path+"/"+configexcelname[i]+".xlsx");
            if(!file.exists()){
//                QString cur="...添加失败，缺少"+configexcelname[i];
//                ui->textEdit_2->append(cur);
                all_event_time->quit();
                delete all_event_time;
//                emit lack();
                event1->quit();
                delete event1;
                log_excelname.clear();
                log_excelname.append(configexcelname[i]);
                log_excelname.append("0");
                return log_excelname;
            }

            if(i==0 && configexcelname[0]=="all_event_time"){ //对all_event_time表特殊处理
                QString ename=event1->getCellValue(row,3).toString();
                QString econfig=event1->getCellValue(row,4).toString();//获取配置

                QString cur_time=event1->add_time(starttime,eventtype_timelist[eventtypelist.indexOf(eventtype)]*5);
                all_event_time->seteventconfig(eventtype,ename,econfig,cur_time);//添加all_event_time配置

                eventtype_timelist[eventtypelist.indexOf(eventtype)]++;

                if(log_excelname.indexOf(configexcelname[i])==-1){
//                    ui->textEdit_2->insertPlainText("\""+configexcelname[i]+"\"");
                    log_excelname.append(configexcelname[i]);
                }
                continue;
            }

            bool isupdate=copycolumn_excel(eventtype,row+i,3,configexcelname[i],1,30);

            if(log_excelname.indexOf(configexcelname[i])==-1 && isupdate){
//                ui->textEdit_2->insertPlainText("\""+configexcelname[i]+"\"");
                log_excelname.append(configexcelname[i]);
            }
        }
    }

    all_event_time->quit();
    delete all_event_time;
//    emit complete();
    event1->quit();
    delete event1;
    return log_excelname;
}

bool ConfigTab::copycolumn_excel(QString sheet,int row,int col,QString excelname,int rows,int cols)//表格之间复制一定区域
{
    QExcel *event1=new QExcel();
    event1->open(path+"/配置汇总.xlsx");

    QExcel* config_update=new QExcel();
    config_update->open(path+"/"+excelname+".xlsx");

    if(sheet!="副本"&&sheet!="其余功能"){
        sheet="主线活动";
    }
    event1->selectSheet(sheet);
    QAxObject* Range_1 = event1->sheet->querySubObject ("Cells(int,int)",row,col);
    Range_1 = Range_1->querySubObject("Resize(int,int)",rows,cols);
    QVariant var_1 = Range_1->dynamicCall("Value");


    int row2=config_update->getUsedRowsCount();
    QList< QList<QVariant> >  datas;//构建数据模型用于存储数据
    config_update->Qvariant2listlistVariant(var_1,datas);
    for (int i=2;i<=row2;i++) {
        QVariant cur=config_update->getCellValue(i,1).toString();
        if(cur==datas[0][0]){
            config_update->quit();
            delete config_update;
            event1->quit();
            delete event1;
            return false;
        }

    }

    Range_1 = config_update->sheet->querySubObject ("Cells(int,int)",row2+1,1);
    Range_1 = Range_1->querySubObject("Resize(int,int)",rows,cols);
    Range_1->setProperty ("Value",var_1);

    config_update->save();
    config_update->quit();
    delete config_update;
    event1->quit();
    delete event1;
    return true;

}

