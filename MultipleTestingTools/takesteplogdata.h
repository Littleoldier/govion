#ifndef TAKESTEPLOGDATA_H
#define TAKESTEPLOGDATA_H

#include <QWidget>
#include <QDir>
#include <synchapi.h>
#include <QFile>
#include <qdebug.h>
#include <QFileDialog>
#include <QStringListModel>
#include <QDateTime>
#include <QMessageBox>
#include <QAxObject>
#include<ole2.h>
#include "recordlog.h"
#include <QThread>
namespace Ui {
class TakeStepLogData;
}

class TakeStepLogData : public QWidget
{
    Q_OBJECT

public:
    explicit TakeStepLogData(QWidget *parent = nullptr);
    ~TakeStepLogData();

    void showdata(QString data);
    bool ResultDataAnalyze();

    void initlog();
signals:
    void Startsignal();
    void CloseTakeStepLogData();//关闭窗口

protected:
    void closeEvent(QCloseEvent *event);

private slots:
    void on_pushButton_2_clicked();

    void on_pushButton_clicked();

private:
    Ui::TakeStepLogData *ui;

    RecordLog *recrdlog;
    QThread *Sub_Thread;

    QDir *dir ;
    QString filestrPath ;

    QStringList filePath;
    QStringList fileName;
    QStringList fileTime;
    QStringList TebleSpend;

};

#endif // TAKESTEPLOGDATA_H
