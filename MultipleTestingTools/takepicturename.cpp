#include "takepicturename.h"
#include "ui_takepicturename.h"

TakePictureName::TakePictureName(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::TakePictureName)
{
    ui->setupUi(this);
    this->setFixedSize(993,311);
    recrdlog = new RecordLog;//动态分配空间，不能指定父对象
    Sub_Thread = new QThread(this);//创建子线程
    recrdlog->moveToThread(Sub_Thread);//把自定义的线程加入子线程
    connect(this,SIGNAL(Startsignal()),recrdlog,SLOT(ThreadLog()));//日志线程
    connect(Sub_Thread, SIGNAL(finished()), recrdlog, SLOT(deleteLater())); //释放线程

    Sub_Thread->start(); //启动子线程
    emit Startsignal();
}

TakePictureName::~TakePictureName()
{
    Sub_Thread->quit();
    Sub_Thread->wait();//线程退出--不可delete线程对象
    delete ui;
}


bool isDigitString(const QString& src) {
    const char *s = src.toUtf8().data();
    while(*s && *s>='0' && *s<='9')s++;
    return !bool(*s);
}

//获取文件夹路径
void TakePictureName::on_pushButton_clicked()
{
    QString path = QFileDialog::getExistingDirectory(this, tr("选择文件夹"),"D:\\",QFileDialog::ShowDirsOnly);
    if(path.isEmpty()){
        return;
    }
    ui->PathlineEdit->setText(path);
    dir = new QDir(path);
    int p = path.lastIndexOf("/");
    path = path.mid(p+1);
    filestr = path;
    qDebug()<<path;

    fileDir.clear();
    fileName.clear();
    TebleDir.clear();
    TebleName.clear();
    ui->lineEdit->clear();
    ui->lineEdit_2->clear();
    initlog();

    QStringList filter;
    QList<QFileInfo> *fileInfo=new QList<QFileInfo>(dir->entryInfoList(filter));
    for(int i = 0;i<fileInfo->count(); i++)
    {
        if("." == fileInfo->at(i).fileName() || ".." == fileInfo->at(i).fileName())//过滤.和..隐藏文件
            continue;
        if(fileInfo->at(i).isDir()){//判断是否为文件夹

            if(isDigitString(fileInfo->at(i).fileName())){//判断是否为纯数字

                sub_dir = new QDir(fileInfo->at(i).filePath()+"/photo");

                fileDir.append(fileInfo->at(i).fileName());//添加文件夹名
                strname.clear();//清空文件名列表

                QStringList sub_filter;
                QList<QFileInfo> *sub_fileInfo=new QList<QFileInfo>(sub_dir->entryInfoList(sub_filter));
                for(int j = 0;j<sub_fileInfo->count(); j++)
                {
                    if("." == sub_fileInfo->at(j).fileName() || ".." == sub_fileInfo->at(j).fileName())//过滤.和..隐藏文件
                        continue;
                    if(sub_fileInfo->at(j).isFile()){
                        if(sub_fileInfo->at(j).fileName().contains("_"))
                        {
                            int point = sub_fileInfo->at(j).fileName().lastIndexOf("_");
                            QString str = sub_fileInfo->at(j).fileName().mid(point+1);
                            int p = str.lastIndexOf(".");
                            str = str.mid(0,p);
                            strname += str+",";//添加文件名
                        }
                    }
                }
                strname = strname.left(strname.size()-1);
                fileName.append(strname);
            }
        }
    }

    if(!fileDir.isEmpty()){

        for(int i = 1; i <fileDir.count()+2; i++){ //排序文件夹
            for (int j = 0 ; j < fileDir.count(); j++) {
                if(QString::number(i) == fileDir.at(j)){
                    logRecord(fileDir.at(j)+" "+fileName.at(j));
                    TebleDir.append(fileDir.at(j));
                    TebleName.append(fileName.at(j));
                }
            }
        }

        QString start = TebleDir.at(0);
        QString endstr = TebleDir.at(TebleDir.count()-1);

        logRecord(filestr+" 已读文件数 "+QString::number(TebleDir.count()));
        ui->lineEdit_2->setText(start+"-"+endstr);
        int na;
        int cp = 0;
        QString nameberstr;
        for (int j = 0 ; j < TebleDir.count(); j++) { //查找缺失文件
            na = j + 1;
            if(QString::number(na + cp) != TebleDir.at(j)){
                qDebug()<<na + cp;
                qDebug()<<TebleDir.at(j);
                nameberstr += QString::number(na+cp) +",";
                ui->lineEdit->setText(nameberstr);
                logRecord("缺失文件夹 "+QString::number(na+cp));
                cp = cp + 1;
            }
        }
        logRecord("读取数据完成");
    }else {
        logRecord("请选择正确的文件路径");
    }
}

//写入Ecxel表
void TakePictureName::on_pushButton_2_clicked()
{
    if(TebleDir.isEmpty()){
        return;
    }
    logRecord("准备写入");//提示
    QDir *DataFile = new QDir;
    bool exist = DataFile->exists("PictureExcel");
    if(!exist){
        bool isok = DataFile->mkdir("PictureExcel"); // 新建文件夹
            if(!isok){
               logRecord("创建文件夹Excel失败") ;//提示
               recrdlog->write();
                return;
             }
    }

    QDateTime current_date_time =QDateTime::currentDateTime();
    QString current_date =current_date_time.toString("yyyyMMdd_hhmmss");
    filestrPath = QCoreApplication::applicationDirPath();//获取当前程序路径
    QString absFilePath =filestrPath + "/PictureExcel/"+current_date+"_"+filestr+".xlsx";
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
     cellrowlist<<"A"<<"B";
     QString str;
      qDebug()<<"设置标题";
     //设置标题
     for(int i = 0; i < cellrowlist.count(); i++){
        QString str = cellrowlist.at(i)+QString::number(1);
        QString Tostr;
        if(i == 0) Tostr = "文件夹";
        else Tostr = "图名集";

       QAxObject * cell = worksheet->querySubObject("Range(QVariant, QVariant)",str);//获取单元格
        cell->dynamicCall("SetValue(const QVariant&)",QVariant(Tostr));//设置单元格的值
        cell->setProperty("HorizontalAlignment", -4108); //左对齐（xlLeft）：-4131  居中（xlCenter）：-4108  右对齐（xlRight）：-415

     }
     logRecord("开始写入");//提示

     for (int i = 0 ; i < TebleDir.count(); i++) {
        for(int j = 0; j <cellrowlist.count(); j++) {
            counstr = cellrowlist.at(j)+QString::number(i+2);
            Colorcell = worksheet->querySubObject("Range(QVariant, QVariant)",counstr);//获取单元格
            if(j == 0){
                Colorcell->dynamicCall("SetValue(const QVariant&)",QVariant(TebleDir.at(i)));//设置单元格的值
            }
            else{
                 Colorcell->dynamicCall("SetValue(const QVariant&)",QVariant(TebleName.at(i)));//设置单元格的值
            }
            Colorcell->setProperty("HorizontalAlignment", -4108); //左对齐（xlLeft）：-4131  居中（xlCenter）：-4108  右对齐（xlRight）：-4152
        }
     }
     workbook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(absFilePath)); //保存到filepath
     workbook->dynamicCall("Close (Boolean)", false);  //关闭文件
     excel->dynamicCall("Quit(void)");  //退出
     logRecord("输出文件："+absFilePath) ;//提示
     recrdlog->write();
     delete excel;
     excel = nullptr;

}

void TakePictureName::logRecord(QString data)
{
    QDateTime da_time;
    QString time_str = da_time.currentDateTime().toString("yyyy-MM-dd HH:mm:ss");
    QString logstr = "["+time_str+"]" +data +"\n";;
    ui->textEdit->append(logstr);
    recrdlog->strlog += logstr;
}

void TakePictureName::closeEvent(QCloseEvent *event)
{
    recrdlog->quit();
    emit CloseTakePictureName();
    qDebug()<<"send CloseTakePictureName ok";

}

//初始化log
void TakePictureName::initlog()
{
   if(false == recrdlog->bInitLog){
        recrdlog->init("TakePictureNameLog");
   }
}
