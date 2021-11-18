#ifndef TAKEPICTURENAME_H
#define TAKEPICTURENAME_H

#include <QWidget>
#include <QDir>
#include <QFile>
#include <qdebug.h>
#include <QFileDialog>
#include <QStringListModel>
#include <QDateTime>
#include <QMessageBox>
#include <QAxObject>
#include "recordlog.h"
#include <QThread>
namespace Ui {
class TakePictureName;
}

class TakePictureName : public QWidget
{
    Q_OBJECT

public:
    explicit TakePictureName(QWidget *parent = nullptr);
    ~TakePictureName();
    void logRecord(QString data);

    void initlog();
signals:
    void Startsignal();
    void CloseTakePictureName();//关闭窗口

protected:
    void closeEvent(QCloseEvent *event);
private slots:
    void on_pushButton_clicked();

    void on_pushButton_2_clicked();

private:
    Ui::TakePictureName *ui;

    RecordLog *recrdlog;
    QThread *Sub_Thread;

    QDir *dir ;
    QDir *sub_dir ;
    QString filestrPath ;

    QStringList filePath;
    QStringList fileName;
    QStringList fileDir;
    QStringList TebleDir;
    QStringList TebleName;

    QString strname;
    QString filestr;
};

#endif // TAKEPICTURENAME_H
