#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QCoreApplication>
#include <QDebug>
#include <QFile>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QJsonValue>
#include <QStringList>
#include <QTextStream>
#include <QProcess>
#include <QFileInfo>
#include <QDir>
#include <QDebug>
#include <exception>
#include <QAxObject>
#include <QDateTime>
#include <QFileDialog>
#include <QDesktopServices>
#include <QTextCodec>
#include <QSettings>
#include <QtAlgorithms>
#include "QDebugOutputRedirector.h"
QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

struct TwoInts {
    int first;
    int second;
};
class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    void compareArrays(const QJsonArray &array1, const QJsonArray &array2, const QString &path, QTextStream &outputStream);
    void compareJSONArrays(const QJsonArray &array1, const QJsonArray &array2, const QString &path, QTextStream &outputStream);
    void compareJSONFiles(const QString &filePath1, const QString &filePath2, const QString &outputFilePath);
    //    void myMessageHandler(QtMsgType type, const QMessageLogContext &context, const QString &msg);

    QString filePath ;
    QString DatafilePath ;

    int InseData;
    int IOutData;
    int listcount;
    int rowcount;
    QString StrData;

    //*********************设置配置
    QDir *Setdir ;

    QStringList Set_MainKeylist;
    QStringList Set_InformationArrayKeylist;
    QStringList Set_AttachInfoKeylist;
    QStringList Set_RotatedRectKeylist;

    QStringList ExcelKeylist;

    QStringList listcountlist;
    QStringList rowcountlist;
    QStringList counstrlist;

    QStringList listcountlist_1;
    QStringList rowcountlist_1;
    QStringList counstrlist_1;

    // 存储所有值的链表
    QStringList valuesList;
    QStringList valuesKeyList;
    QMap<QString, QString> map;


    QString folderPath ;
    QString folderPathA ;
    QString folderPathB ;

    QString EcxelFilePath ;
    QAxObject *lookexcel;

    // 读取上一次选择的路径
    QSettings settings_only;
    QSettings settings_1;
    QSettings settings_2;

    QJsonObject config ;

    QString jsonpath;
    int ModecomboBox;
    QString comboBoxIndex;
    QString comboBoxIndex_2;

    void RecordDataToExcel(QString Path);
    void setExcelTebel();
    void setData(int Setlistcount, int Setrowcount, QString SetStrData);
    void printJsonValueType(const QJsonValue &valueA, const QJsonValue &valueB, QString keys, QTextStream &diffStream, QString pathA, QString pathB,int count);
    void compareJson(const QJsonObject &objA, const QJsonObject &objB, const QString &pathA, const QString &pathB, QTextStream &diffStream,int count);
    void compareJsonFiles(const QString &filePathA, const QString &filePathB, QTextStream &diffStream,QString fileA,QString fileB,int count);
    void processFolder(const QString &folderPathA, const QString &folderPathB, QTextStream &diffStream,bool &MissingAttributeBool);
    void setDifferencequantityData(int Setrowcount, int SetDifferencequantity);
    void DifferencequantityAppendlist();
    void processFolder_1(const QString &folderPathA, QTextStream &diffStream);
    void openFolder(const QString &folderPath);

    void ClearAllDatalist();
    void readjsonfile();
    void setData(QString SetKey, int Setrowcount, QString SetStrData);
    QPair<QJsonObject, QJsonObject> getJsonObjects(const QJsonObject &objA, const QJsonObject &objB);
    QJsonObject ProcessingObjects(const QJsonObject &Object);
    void getSetConfigureFieldsObjects();
    void removeFieldsFromJson(QJsonObject& jsonObject, const QStringList &fieldsToRemove);
    void readAllKeys(const QJsonObject& jsonObjectA, const QJsonObject& jsonObjectB,QStringList& keyListA,QStringList& keyListB);
    void getAllKeys(const QJsonObject &jsonObject, QStringList &keyList);
    QStringList findDifferentValues(const QStringList &list1, const QStringList &list2);
    void ObtainFilesWithTheSameStructure(const QStringList &subfoldersA, const QStringList &subfoldersB, QString &fileNameA, QString &fileNameB);
    void CheckAttributes(const QStringList &filePathA, const QStringList &filePathB, const QString& folderPathA, const QString& folderPathB,bool &MissingAttributeBool);
    void CheckJsonFiles(const QString &filePathA, const QString &filePathB,bool &MissingAttributeBool);
    QStringList findDifferentFields(const QStringList &list1, const QStringList &list2);
    void CheckFileNameFormat(const QString &folderPathA, const QString &folderPathB,bool& BoolFormatFlag);
    void CheckFileNameFormat_1(const QString &folderPathA, bool &BoolFormatFlag);
    void printJsonValueType_B(const QJsonValue &valueA, const QJsonValue &valueB, QString keys, QTextStream &diffStream, QString pathA, QString pathB, int count);
    void handleCleanup(QAxObject* workbook);
    void checkResultLJson(QString path);
private slots:
    void on_pushButton_clicked();

    void on_pushButton_2_clicked();

    void on_pushButton_3_clicked();

    void on_pushButton_4_clicked();

    void on_pushButton_5_clicked();

    void on_comboBox_currentIndexChanged(int index);

    void on_comboBox_2_currentIndexChanged(int index);

    void on_pushButton_6_clicked();

private:
    Ui::MainWindow *ui;
};
#endif // MAINWINDOW_H
