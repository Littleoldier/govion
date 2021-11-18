#ifndef BATCHDELETIONDOCUMENT_H
#define BATCHDELETIONDOCUMENT_H

#include <QWidget>
#include <QDir>
#include <QFile>
#include <qdebug.h>
#include <QFileDialog>
#include <QStringListModel>
#include <QDateTime>
#include <QMessageBox>
#include "recordlog.h"
#include <QThread>

namespace Ui {
class BatchDeletionDocument;
}

class BatchDeletionDocument : public QWidget
{
    Q_OBJECT

public:
    explicit BatchDeletionDocument(QWidget *parent = nullptr);
    ~BatchDeletionDocument();

    void definitionReset();
    void writeData(QString strtip);
    void logRecord(QString log);

    void initlog();
signals:
    void Startsignal();
    void CloseBatchDeletionDocument();//关闭窗口


protected:
    void closeEvent(QCloseEvent *event);
private slots:
    void on_pushButton_clicked();

    void on_filelistView_clicked(const QModelIndex &index);

    void on_deletefilelistView_2_clicked(const QModelIndex &index);

    void on_pushButton_2_clicked();

private:
    Ui::BatchDeletionDocument *ui;

    RecordLog *recrdlog;
    QThread *Sub_Thread;

    QDir *dir ;
    QDir *sub_dir ;

    QStringList filePath;
    QStringList fileName;
    QStringList Removestrlist;
    QStringList recordRemovelist;
    QStringListModel *model;

    int filecount;
    int filetypecount;
    int choosefilecount;
    int Removefilecount;
    int NGfilecount;

};

#endif // BATCHDELETIONDOCUMENT_H
