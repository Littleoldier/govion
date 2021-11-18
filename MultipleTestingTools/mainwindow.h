#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include "batchdeletiondocument.h"
#include "resultdataanalyze.h"
#include "takepicturename.h"
#include "takesteplogdata.h"
#include "recordlog.h"
#include <QThread>
namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    void logRecord(QString log);
    void initlog();
signals:
    void Startsignal();
public slots:
    void CloseResultDataAnlyze();
    void CloseBatchDeletionDocument();
    void CloseTakePictureName();
    void CloseTakeStepLogData();

protected:
    void closeEvent(QCloseEvent *event);

private slots:
    void on_ResultDataButton_clicked();

    void on_BatchDeletionButton_clicked();

    void on_TakeNameButton_clicked();

    void on_TakeStepDataButton_2_clicked();

private:
    Ui::MainWindow *ui;
    RecordLog *recrdlog;
    QThread *Sub_Thread;

    BatchDeletionDocument *deletionDialog;
    ResultDataAnalyze     *ResultDataDialog;
    TakePictureName       *TakenameDialog;
    TakeStepLogData       *TakestepDialog;
};

#endif // MAINWINDOW_H
