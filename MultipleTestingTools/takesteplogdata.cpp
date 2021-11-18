#include "takesteplogdata.h"
#include "ui_takesteplogdata.h"

TakeStepLogData::TakeStepLogData(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::TakeStepLogData)
{
    ui->setupUi(this);
    this->setFixedSize(987,355);

    recrdlog = new RecordLog;//动态分配空间，不能指定父对象
    Sub_Thread = new QThread(this);//创建子线程
    recrdlog->moveToThread(Sub_Thread);//把自定义的线程加入子线程
    connect(this,SIGNAL(Startsignal()),recrdlog,SLOT(ThreadLog()));//日志线程
    connect(Sub_Thread, SIGNAL(finished()), recrdlog, SLOT(deleteLater())); //释放线程

    Sub_Thread->start(); //启动子线程
    emit Startsignal();
}

TakeStepLogData::~TakeStepLogData()
{
    Sub_Thread->quit();
    Sub_Thread->wait();//线程退出--不可delete线程对象
    delete ui;
}

//选择文件
void TakeStepLogData::on_pushButton_clicked()
{
    QString path = QFileDialog::getExistingDirectory(this, tr("选择文件夹"),"D:\\",QFileDialog::ShowDirsOnly);
    if(path.isEmpty()){
        return;
    }
    dir = new QDir(path);
    if("LOG"!=dir->dirName() )
    {
        showdata("请选择正确的路径");
        return;
    }
    ui->PathlineEdit->setText(path);

    initlog();
    showdata("开始读取");

    filePath.clear();
    fileName.clear();
    fileTime.clear();

    QStringList filter;
    QList<QFileInfo> *fileInfo=new QList<QFileInfo>(dir->entryInfoList(filter));
    for(int i = 0;i<fileInfo->count(); i++)
    {
        if("." == fileInfo->at(i).fileName() || ".." == fileInfo->at(i).fileName())//过滤.和..隐藏文件
            continue;
        filePath.append(fileInfo->at(i).filePath()+"/stepLog.log");
        fileName.append(fileInfo->at(i).fileName());
        qDebug()<<fileInfo->at(i).lastRead().toString();
        fileTime.append(fileInfo->at(i).lastRead().toString("yyyy/MM/dd hh:mm:ss"));

    }
    if(ResultDataAnalyze()){
        showdata("读取完成");
    }else {
        showdata("读取失败");
        recrdlog->write();
    }
}

bool TakeStepLogData::ResultDataAnalyze()
{
    for(int i = 0 ; i < filePath.count(); i++){//读取LOG文件
        showdata(filePath.at(i));

        QFile file(filePath.at(i));//打开当前Result.json文件
        if(!file.open(QIODevice::ReadOnly |QIODevice::Text) ){//以读的方式打开文件
            showdata(filePath.at(i)+" "+"file.errorString()");
            TebleSpend.clear();
            return false;
        }

        QTextStream in(&file);
        in.setCodec("GBK"); // 设置文件的编码格式为GBK
        QString temp = in.readAll();
        if(temp.contains("[*****等待算法结果*****]")){
            QStringList sit = temp.split("[*****等待算法结果*****]");
            if(sit.at(1).contains("Spend")){
                QStringList list = sit.at(1).split("Spend:");
                if(list.at(1).contains("[")){
                    QStringList kli = list.at(1).split("[");
                    TebleSpend.append(kli.at(0).trimmed());
                }
            }
        }
   }
    return  true;
}

//写入数据
void TakeStepLogData::on_pushButton_2_clicked()
{
    if(TebleSpend.isEmpty())
    {
        return;
    }
    showdata("准备写入");//提示
    QDir *DataFile = new QDir;
    bool exist = DataFile->exists("SteplogExcel");
    if(!exist)
    {
        bool isok = DataFile->mkdir("SteplogExcel"); // 新建文件夹
            if(!isok){
               showdata("创建文件夹SteplogExcel失败") ;//提示
               recrdlog->write();
                return;
             }
    }

    QDateTime current_date_time =QDateTime::currentDateTime();
    QString current_date =current_date_time.toString("yyyyMMdd_hhmmss");
    filestrPath = QCoreApplication::applicationDirPath();//获取当前程序路径
    QString absFilePath =filestrPath + "/SteplogExcel/"+current_date+"_"+".xlsx";
    absFilePath = absFilePath.replace("/", "\\");
     QAxObject *excel = new QAxObject(this);//建立excel操作对象
     excel->setControl("Excel.Application");//连接Excel控件
     excel->dynamicCall("SetVisible (bool Visible)","false");//设置为不显示窗体
     excel->setProperty("DisplayAlerts", false);//不显示任何警告信息，如关闭时的是否保存提示

     //创建ex表
     QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
     workbooks->dynamicCall("Add");//新建一个工作簿
     QAxObject *workbook = excel->querySubObject("ActiveWorkBook");//获取当前工作簿
     QAxObject *worksheets = workbook->querySubObject("Sheets");//获取工作表集合
     QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);//获取工作表集合的工作表1，即sheet1
     QAxObject *Colorcell;
     QVariant var;
     QString counstr;

     QStringList cellrowlist;
     cellrowlist<<"A"<<"B"<<"C";
     QString str;
      qDebug()<<"设置标题";
     //设置标题
     for(int i = 0; i < cellrowlist.count(); i++){
        QString str = cellrowlist.at(i)+QString::number(1);
        QString Tostr;
        if(i == 0) Tostr = "文件夹";
        else if(i == 1) Tostr = "文件夹时间";
        else if(i == 2) Tostr = "Spend";

       QAxObject * cell = worksheet->querySubObject("Range(QVariant, QVariant)",str);//获取单元格
        cell->dynamicCall("SetValue(const QVariant&)",QVariant(Tostr));//设置单元格的值
        cell->setProperty("HorizontalAlignment", -4108); //左对齐（xlLeft）：-4131  居中（xlCenter）：-4108  右对齐（xlRight）：-415

     }
     showdata("开始写入");//提示

     for (int i = 0 ; i < fileName.count(); i++) {
        for(int j = 0; j <cellrowlist.count(); j++) {
            counstr = cellrowlist.at(j)+QString::number(i+2);
            Colorcell = worksheet->querySubObject("Range(QVariant, QVariant)",counstr);//获取单元格
            if(j == 0){
                Colorcell->dynamicCall("SetValue(const QVariant&)",QVariant(fileName.at(i)));//设置单元格的值
            }
            else if(j == 1){
                 Colorcell->dynamicCall("SetValue(const QVariant&)",QVariant(fileTime.at(i)));//设置单元格的值
            }
            else if(j == 2){
                 Colorcell->dynamicCall("SetValue(const QVariant&)",QVariant(TebleSpend.at(i)));//设置单元格的值
            }
            Colorcell->setProperty("HorizontalAlignment", -4108); //左对齐（xlLeft）：-4131  居中（xlCenter）：-4108  右对齐（xlRight）：-4152
        }
     }
     workbook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(absFilePath)); //保存到filepath
     workbook->dynamicCall("Close (Boolean)", false);  //关闭文件
     excel->dynamicCall("Quit(void)");  //退出
     showdata("输出文件："+absFilePath) ;//提示
     recrdlog->write();
     delete excel;
     excel = nullptr;
}


void TakeStepLogData::showdata(QString data)
{
    QString logstr ;
    QDateTime da_time;
    QString time_str = da_time.currentDateTime().toString("yyyy-MM-dd HH:mm:ss");
    logstr = "["+time_str+"]" +data +"\n";;
    ui->textEdit->append(logstr);
    recrdlog->strlog += logstr;
}


void TakeStepLogData::closeEvent(QCloseEvent *event)
{
    recrdlog->quit();
    emit CloseTakeStepLogData();
    qDebug()<<"send CloseTakeStepLogData ok";

}

//初始化log
void TakeStepLogData::initlog()
{
   if(false == recrdlog->bInitLog){
        recrdlog->init("TakeStepLogDataLog");
   }
}

