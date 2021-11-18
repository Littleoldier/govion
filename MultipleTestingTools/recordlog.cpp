#include "recordlog.h"

RecordLog::RecordLog(QObject *parent):
    QObject(parent)
{
    bInitLog = false;
    bWritelog = false;
    strlog = "";
}
RecordLog::~RecordLog()
{
    if(file != nullptr)
        delete file;
    file = nullptr;
}

bool RecordLog::InitRecordLog()
{
    QDateTime da_time;
    QString time_str = da_time.currentDateTime().toString("yyyyMMddHHmmss");

    QDir *DataFile = new QDir;
    bool exist = DataFile->exists(logname);
    if(!exist)
    {
        bool isok = DataFile->mkdir(logname); // 新建文件夹
            if(!isok)
                    return false;
    }
    QString fileName = logname+"/"+time_str+"LogData.txt";

    file = new QFile(fileName);

    if(!file->open(QIODevice::WriteOnly|QIODevice::Text|QIODevice::Append))
    {
        return false;
    }
    return true;
}

//写入数据
bool RecordLog::writeData(QFile *file)
{
    QTextStream stream(file);
    stream<<strlog;
    file->close();
    strlog.clear();//清空数据
    return true;
}

//指令
void RecordLog::ThreadLog()
{
    qDebug()<<"startTread";

    while(1){
        if("init" == Order){
            qDebug()<<"Order ==  init";
            bInitLog  = InitRecordLog();//初始化
            mutex.lock();
            Order.clear();
            mutex.unlock();
        }
        if("Close"== Order){
            qDebug()<<"Order ==  Close";
            mutex.lock();
            bWritelog = writeData(file);//写入
            qDebug()<<"bWritelog ="<<bWritelog;
            if(bWritelog){
                if(file != nullptr)
                    delete file;
                file = nullptr;
                bInitLog = false;
            }
            Order.clear();
            mutex.unlock();
        }
        if("Quit" == Order){
            qDebug()<<"Order ==  Quit";
            mutex.lock();
            Order.clear();
            mutex.unlock();
            break;
        }

        if("Close and Quit"== Order ){
            qDebug()<<"Order ==  Close and Quit";
            mutex.lock();
            bWritelog = writeData(file);//写入
            qDebug()<<"bWritelog ="<<bWritelog;
            if(bWritelog){
                if(file != nullptr)
                    delete file;
                file = nullptr;
                bInitLog = false;
            }
            Order.clear();
            mutex.unlock();
            break;
        }
    }
}

void RecordLog::init(QString log)
{
    mutex.lock();
    Order = "init";
    logname = log;
    mutex.unlock();
}

void RecordLog::write()
{
    mutex.lock();
    Order = "Close";
    mutex.unlock();
}

void RecordLog::quit()
{
    mutex.lock();
    Order = "Quit";
    mutex.unlock();
}

void RecordLog::Writeandquit()
{
    mutex.lock();
    Order = "Close and Quit";
    mutex.unlock();
}
