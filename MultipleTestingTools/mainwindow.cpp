#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    this->setFixedSize(504,326);
    ResultDataDialog = nullptr;
    deletionDialog = nullptr;
    TakenameDialog = nullptr;
    TakestepDialog = nullptr;

    recrdlog = new RecordLog;//动态分配空间，不能指定父对象
    Sub_Thread = new QThread(this);//创建子线程
    recrdlog->moveToThread(Sub_Thread);//把自定义的线程加入子线程
    connect(this,SIGNAL(Startsignal()),recrdlog,SLOT(ThreadLog()));//日志线程
    connect(Sub_Thread, SIGNAL(finished()), recrdlog, SLOT(deleteLater())); //释放线程

    Sub_Thread->start(); //启动子线程
    emit Startsignal();
    initlog();
}

MainWindow::~MainWindow()
{
    Sub_Thread->quit();
    Sub_Thread->wait();//线程退出--不可delete线程对象
    delete ui;
}

void MainWindow::on_ResultDataButton_clicked()
{
    if(nullptr == ResultDataDialog){
        ResultDataDialog = new ResultDataAnalyze();
        ResultDataDialog->show();
        logRecord("开启对比数据程序");
        connect(ResultDataDialog,SIGNAL(CloseResultDataAnlyze()),this,SLOT(CloseResultDataAnlyze()));
    }
}

void MainWindow::on_BatchDeletionButton_clicked()
{
    if(nullptr == deletionDialog){
        deletionDialog = new BatchDeletionDocument();
        deletionDialog->show();
        logRecord("开启批量删除程序");
        connect(deletionDialog,SIGNAL(CloseBatchDeletionDocument()),this,SLOT(CloseBatchDeletionDocument()));
    }
}

void MainWindow::on_TakeNameButton_clicked()
{
    if(nullptr == TakenameDialog){
        TakenameDialog = new TakePictureName();
        TakenameDialog->show();
        logRecord("开启批量读图程序");
        connect(TakenameDialog,SIGNAL(CloseTakePictureName()),this,SLOT(CloseTakePictureName()));
    }

}


void MainWindow::on_TakeStepDataButton_2_clicked()
{
    if(nullptr == TakestepDialog){
        TakestepDialog = new TakeStepLogData();
        TakestepDialog->show();
        logRecord("开启批量获取数据程序");
        connect(TakestepDialog,SIGNAL(CloseTakeStepLogData()),this,SLOT(CloseTakeStepLogData()));
    }
}


void MainWindow::CloseResultDataAnlyze()
{
    ResultDataDialog = nullptr;
    qDebug()<<"delete  ResultDataDialog ok";
    logRecord("关闭对比数据程序");
}

void MainWindow::CloseBatchDeletionDocument()
{
    deletionDialog = nullptr;
    qDebug()<<"delete  deletionDialog ok";
    logRecord("关闭批量删除程序");
}

void MainWindow::CloseTakePictureName()
{
    TakenameDialog = nullptr;
    qDebug()<<"delete  TakenameDialog ok";
    logRecord("关闭批量读图程序");
}

void MainWindow::CloseTakeStepLogData()
{
    TakestepDialog = nullptr;
    qDebug()<<"delete  TakestepDialog ok";
    logRecord("关闭批量获取数据程序");
}

void MainWindow::closeEvent(QCloseEvent *event)
{
    if(nullptr != ResultDataDialog ||nullptr != deletionDialog
      ||nullptr != TakenameDialog ||nullptr != TakestepDialog){

        event->ignore();

    }else {
       recrdlog->Writeandquit();
    }
}


void MainWindow::logRecord(QString log)
{
    QDateTime da_time;
    QString time_str = da_time.currentDateTime().toString("yyyy-MM-dd HH:mm:ss");
    QString logstr = "["+time_str+"]" +log +"\n";;
    ui->textEdit->append(logstr);
    recrdlog->strlog += logstr;
}

//初始化log
void MainWindow::initlog()
{
   if(false == recrdlog->bInitLog){
        recrdlog->init("MultipleTestingToolsLog");
   }
}

