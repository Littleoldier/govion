#include "batchdeletiondocument.h"
#include "ui_batchdeletiondocument.h"

BatchDeletionDocument::BatchDeletionDocument(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::BatchDeletionDocument)
{
    ui->setupUi(this);
    this->setFixedSize(1003,616);
    recrdlog = new RecordLog;//动态分配空间，不能指定父对象
    Sub_Thread = new QThread(this);//创建子线程
    recrdlog->moveToThread(Sub_Thread);//把自定义的线程加入子线程
    connect(this,SIGNAL(Startsignal()),recrdlog,SLOT(ThreadLog()));//日志线程
    connect(Sub_Thread, SIGNAL(finished()), recrdlog, SLOT(deleteLater())); //释放线程

    Sub_Thread->start(); //启动子线程
    emit Startsignal();
}

BatchDeletionDocument::~BatchDeletionDocument()
{
    Sub_Thread->quit();
    Sub_Thread->wait();//线程退出--不可delete线程对象
    delete ui;
}

void BatchDeletionDocument::on_pushButton_clicked()
{
    definitionReset();//Reset
    QString path = QFileDialog::getExistingDirectory(this, tr("选择文件夹"),"D:\\",QFileDialog::ShowDirsOnly);
    if(path.isEmpty()){
        return;
    }

    ui->PathlineEdit->setText(path);
    dir = new QDir(path);

    QStringList filter;
    QList<QFileInfo> *fileInfo=new QList<QFileInfo>(dir->entryInfoList(filter));
    for(int i = 0;i<fileInfo->count(); i++)
    {
        if("." == fileInfo->at(i).fileName() || ".." == fileInfo->at(i).fileName())//过滤.和..隐藏文件
            continue;
       sub_dir = new QDir(fileInfo->at(i).filePath());

        QStringList sub_filter;
        QList<QFileInfo> *sub_fileInfo=new QList<QFileInfo>(sub_dir->entryInfoList(sub_filter));
        for(int j = 0;j<sub_fileInfo->count(); j++)
        {
            if("." == sub_fileInfo->at(j).fileName() || ".." == sub_fileInfo->at(j).fileName())//过滤.和..隐藏文件
                continue;
            qDebug()<<sub_fileInfo->at(j).fileName();
            filePath.append(sub_fileInfo->at(j).filePath());
            fileName.append(sub_fileInfo->at(j).fileName());
            filecount++;
        }
    }

    for(int i = 0 ; i < fileName.count(); i++){      //去掉重文件名
      for(int j = i + 1; j < fileName.count(); j++){
          if(fileName.at(i) == fileName.at(j)){
              fileName.removeAt(j);
              j--;
          }
      }
    }

    model = new QStringListModel(this);
    model->setStringList(fileName);
    ui->filelistView->setModel(model);

    filetypecount = fileName.count();
    ui->lineEdit->setText(QString::number(filecount));
    ui->lineEdit_2->setText(QString::number(filetypecount));

}

//添加
void BatchDeletionDocument::on_filelistView_clicked(const QModelIndex &index)
{
    if(index.data().toString() == "photo"){
        QMessageBox messagebox(QMessageBox::Warning,"注意","已选中photo文件夹，请检查确认是否批量删除此类文档",QMessageBox::Yes|QMessageBox::No ,NULL);
        if(messagebox.exec() == QMessageBox::No){
            return;
        }
    }
    Removestrlist.push_back(index.data().toString());
    model = new QStringListModel(Removestrlist);
    ui->deletefilelistView_2->setModel(model);

    choosefilecount = Removestrlist.count();
    ui->lineEdit_3->setText(QString::number(choosefilecount));

    for(int i = 0; i < fileName.count(); i++ ){
       if(fileName.at(i) == index.data().toString()){
           fileName.removeAt(i);
       }
    }
    model = new QStringListModel(fileName);
    ui->filelistView->setModel(model);

    filetypecount = fileName.count();
    ui->lineEdit_2->setText(QString::number(filetypecount));

}

//去除已选择文件
void BatchDeletionDocument::on_deletefilelistView_2_clicked(const QModelIndex &index)
{
    for(int i = 0; i < Removestrlist.count(); i++)
    {
        if(Removestrlist.at(i) == index.data().toString() )
        {
            Removestrlist.removeAt(i);
            fileName.append(index.data().toString());
        }
    }
    model = new QStringListModel(Removestrlist);
    ui->deletefilelistView_2->setModel(model);

    model = new QStringListModel(fileName);
    ui->filelistView->setModel(model);

    choosefilecount = Removestrlist.count();
    ui->lineEdit_3->setText(QString::number(choosefilecount));

    filetypecount = fileName.count();
    ui->lineEdit_2->setText(QString::number(filetypecount));

}

void BatchDeletionDocument::definitionReset()
{
    filecount = 0;
    filetypecount = 0;
    choosefilecount = 0;
    Removefilecount = 0;
    NGfilecount = 0;
    fileName.clear();
    filePath.clear();
    Removestrlist.clear();
    recordRemovelist.clear();

    initlog();//初始化日志
}

//开始删除选择文件
void BatchDeletionDocument::on_pushButton_2_clicked()
{
    if(Removestrlist.isEmpty())
        return;

  QMessageBox messagebox(QMessageBox::Warning,"注意","此操作将不可逆转，请检查确认是否批量删除此类文档",QMessageBox::Yes|QMessageBox::No ,NULL);
  if(messagebox.exec() == QMessageBox::Yes){
      logRecord("开始批量删除文件...");
      QDir dir;
      for(int i = 0 ; i < Removestrlist.count(); i++){
          for(int j = 0 ; j < filePath.count(); j++){
              QString strfile = filePath.at(j);
              int nIndex = filePath.at(j).lastIndexOf("/"); // 到下一次出现时
              QString strName = filePath.at(j).mid(nIndex+1,strfile.size());// 从0开始取nIndex个

              if(Removestrlist.at(i) == strName ){
                  bool OkRemovestr = QFile::remove(strfile);
                  if(OkRemovestr){
                      logRecord("Remove OK "+strfile);
                      Removefilecount++;
                      recordRemovelist.append(strName);
                  }else {
                      dir.setPath(strfile);
                      bool bDledir =  dir.removeRecursively();
                      if(bDledir){
                          logRecord("Remove OK "+strfile);
                          Removefilecount++;
                          recordRemovelist.append(strName);
                      }else {
                          logRecord("Remove NG "+strfile);
                          NGfilecount++;
                      }

                  }
              }
          }
      }

      for(int i = 0 ; i < recordRemovelist.count(); i++){      //去掉重文件名
        for(int j = i + 1; j < recordRemovelist.count(); j++){
            if(recordRemovelist.at(i) == recordRemovelist.at(j)){
                recordRemovelist.removeAt(j);
                j--;
            }
        }
      }

      for(int i = 0; i< recordRemovelist.count();i++){
          for(int j = 0; j< Removestrlist.count();j++){
              if(Removestrlist.at(j) == recordRemovelist.at(i)){
                  Removestrlist.removeAt(j);
                  break;
              }
          }
      }
      model = new QStringListModel(Removestrlist);
      ui->deletefilelistView_2->setModel(model);

      choosefilecount = Removestrlist.count();
      ui->lineEdit_3->setText(QString::number(choosefilecount));
      ui->lineEdit_4->setText(QString::number(Removefilecount));
      ui->lineEdit_5->setText(QString::number(NGfilecount));
      logRecord("删除完成！");
      recrdlog->write();
  }
  else { return; }

}

void BatchDeletionDocument::logRecord(QString log)
{
    QDateTime da_time;
    QString time_str = da_time.currentDateTime().toString("yyyy-MM-dd HH:mm:ss");
    QString logstr = "["+time_str+"]" +log +"\n";;
    ui->textEdit->append(logstr);
    recrdlog->strlog += logstr;
}


void BatchDeletionDocument::closeEvent(QCloseEvent *event)
{
    recrdlog->quit();
    emit CloseBatchDeletionDocument();
    qDebug()<<"send CloseBatchDeletionDocument ok";

}

//初始化log
void BatchDeletionDocument::initlog()
{
   if(false == recrdlog->bInitLog){
        recrdlog->init("BatchDeletionDocumentLog");
   }
}
