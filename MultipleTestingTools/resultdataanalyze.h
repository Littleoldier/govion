#ifndef RESULTDATAANALYZE_H
#define RESULTDATAANALYZE_H

#include <QMainWindow>
#include <QWidget>
#include <QByteArray>
#include <QDateTime>
#include <QThread>
#include <QTimer>
#include <QMessageBox>
#include <QListView>
#include <QCloseEvent>
#include <QDir>
#include <QFile>
#include <qdebug.h>
#include <QFileDialog>
#include <QJsonDocument>
#include <QJsonParseError>
#include <QJsonObject>
#include <QJsonArray>
#include <QJsonValue>
#include <QList>
#include <string>
#include <sstream>
#include <iostream>
#include <exception>
#include <QAxObject>
#include"recordlog.h"

namespace Ui {
class ResultDataAnalyze;
}

class ResultDataAnalyze : public QWidget
{
    Q_OBJECT

public:
    explicit ResultDataAnalyze(QWidget *parent = nullptr);
    ~ResultDataAnalyze();

    void setData(int Setlistcount, int Setrowcount, QString SetStrData);

    bool getResultdatas(QStringList &filePathlistlog,QList<QStringList> &mainValue_list,QList<QStringList> &mainKey_list, QVector<QList<QStringList>> &informationArrayValue_vectorlist
         , QVector<QList<QStringList>> &AttachInfoValue_vectorlist, QVector<QList<QStringList>> &RotatedRectValue_vectorlist,QVector<QList<QStringList>> &informationArrayKey_vectorlist,
           QVector<QList<QStringList>> &AttachInfoKey_vectorlist, QVector<QList<QStringList>> &RotatedRectKey_vectorlist);
    void ResultDataCompared();//对比
    void setExcelTebel(QList<QStringList> &mainKey_list,QVector<QList<QStringList>> &informationArrayKey_vectorlist,
                       QVector<QList<QStringList>> &AttachInfoKey_vectorlist, QVector<QList<QStringList>> &RotatedRectKey_vectorlist);//設置表格
    void SetFileTabel();//讀取配置文件

    void RecordDataToExcel(QString filePath);//写入数据到EXCEL表

    void initlog();

    int InseData;
    int IOutData;
    int listcount;
    int rowcount;
    QString StrData;

    void DisPlayTip(QString strTip);
    void sort(QStringList list);
    QString Findfliename(int nlist,QStringList list);
signals:
    void Startsignal();
    void CloseResultDataAnlyze();//关闭窗口

protected:
    void closeEvent(QCloseEvent *event);
private slots:
    void on_PathButton_clicked();

    void on_PathButton_1_clicked();

    void on_StartButton_clicked();

    void on_pushButton_clicked();

    void on_StartButton_2_clicked();

    void on_pushButton_2_clicked();

    void on_inforEdit_textChanged();

private:
    Ui::ResultDataAnalyze *ui;

    RecordLog *recrdlog;
    QThread *Sub_Thread;

    QString filePath ;

    QList<int> nl1;

    //************************ Result *******************//
    //文件1
    int dircount_1;
    QDir *dir ;
    QStringList Filenamelist;
    QStringList FilePathlist;

    QList<QStringList> Q_mainValue_list_1 ;
    QList<QStringList> Q_mainKey_list_1 ;

    QVector<QList<QStringList>> Q_informationArrayValue_vectorlist_1;
    QVector<QList<QStringList>> Q_AttachInfoValue_vectorlist_1;
    QVector<QList<QStringList>> Q_RotatedRectValue_vectorlist_1;

    QVector<QList<QStringList>> Q_informationArrayKey_vectorlist_1;
    QVector<QList<QStringList>> Q_AttachInfoKey_vectorlist_1;
    QVector<QList<QStringList>> Q_RotatedRectKey_vectorlist_1;

    //******************************* Result 2*******************//

    //文件2
    int dircount_2;
    QDir *dir_2 ;
    QStringList Filenamelist_2;
    QStringList FilePathlist_2;

    QList<QStringList> Q_mainValue_list_2 ;
    QList<QStringList> Q_mainKey_list_2;

    QVector<QList<QStringList>> Q_informationArrayValue_vectorlist_2;
    QVector<QList<QStringList>> Q_AttachInfoValue_vectorlist_2;
    QVector<QList<QStringList>> Q_RotatedRectValue_vectorlist_2;

    QVector<QList<QStringList>> Q_informationArrayKey_vectorlist_2;
    QVector<QList<QStringList>> Q_AttachInfoKey_vectorlist_2;
    QVector<QList<QStringList>> Q_RotatedRectKey_vectorlist_2;


   //*********************************************************//
   //******************************* 单个log ******************//

    QDir *dirlog ;
    QStringList Filenamelistlog;
    QStringList FilePathlistlog;

    QStringList LogNamelist;
    QStringList NameCountlist;
    QStringList Namelist;

    //文件1

    QList<QStringList> Q_mainValue_list ;
    QList<QStringList> Q_mainKey_list ;

    QVector<QList<QStringList>> Q_informationArrayValue_vectorlist;
    QVector<QList<QStringList>> Q_AttachInfoValue_vectorlist;
    QVector<QList<QStringList>> Q_RotatedRectValue_vectorlist;

    QVector<QList<QStringList>> Q_informationArrayKey_vectorlist;
    QVector<QList<QStringList>> Q_AttachInfoKey_vectorlist;
    QVector<QList<QStringList>> Q_RotatedRectKey_vectorlist;

    //*********************设置配置
    QDir *Setdir ;
    QString SetFileName;

    QStringList Set_MainKeylist;
    QStringList Set_InformationArrayKeylist;
    QStringList Set_AttachInfoKeylist;
    QStringList Set_RotatedRectKeylist;



    QStringList ExcelKeylist;

    QStringList listcountlist;
    QStringList rowcountlist;
    QStringList counstrlist;

};

#endif // RESULTDATAANALYZE_H
