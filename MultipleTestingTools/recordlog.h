#ifndef RECORDLOG_H
#define RECORDLOG_H

#include <QObject>
#include <QDateTime>
#include <QDir>
#include <QFile>
#include <qdebug.h>
#include <QList>
#include <string>
#include <sstream>
#include <iostream>
#include <exception>
#include <QAxObject>
#include <QMutex>

class RecordLog: public QObject
{
   Q_OBJECT
public:
   explicit RecordLog(QObject *parent = nullptr);
   ~RecordLog();

   bool InitRecordLog();
   bool writeData(QFile *file);

   QString Order;
   QString strlog ;//记录
   QString logname;//log文件夹名

   bool bInitLog;
   bool bWritelog;

   void init(QString logname);
   void write();
   void quit();
   void Writeandquit();
private:
   QFile *file;
   QMutex mutex;

signals:
    void finish();//释放信号
    void DisPlayTipsignal();

public slots:
    void ThreadLog();//线程

};

#endif // RECORDLOG_H
