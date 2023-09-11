#include "mainwindow.h"
#include "ui_mainwindow.h"


MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
     QTextCodec::setCodecForLocale(QTextCodec::codecForName("UTF-8")); // 设置字符编码为UTF-8
    filePath = QApplication::applicationDirPath();
    folderPath.clear();
    folderPathA.clear();
    folderPathB.clear();

    getSetConfigureFieldsObjects();//获取剔除字段配置文件
    jsonpath = "/result.json";
    ModecomboBox = 1;
    EcxelFilePath ="";

    // 获取QPlainTextEdit指针
    QPlainTextEdit* textEdit = ui->plainTextEdit;
    QDebugOutputRedirector redirector(textEdit);
}

MainWindow::~MainWindow()
{
    delete ui;
}


QStringList MainWindow::findDifferentFields(const QStringList& list1, const QStringList& list2)
{
    QStringList uniqueList;

       for (const QString& item : list1) {
           if (!list2.contains(item)) {
               uniqueList.append(item);
           }
       }
    return uniqueList;
}


QStringList MainWindow::findDifferentValues(const QStringList& list1, const QStringList& list2)
{
    QStringList uniqueList;

       for (const QString& item : list1) {
           if (!list2.contains(item)) {
               uniqueList.append(item);
           }
       }

       if(!uniqueList.isEmpty()){
            return uniqueList;
       }
       else {

           for (const QString& item : list2) {
               if (!list1.contains(item)) {
                   uniqueList.append(item);
               }
           }
       }

    return uniqueList;
}


void MainWindow::readAllKeys(const QJsonObject& jsonObjectA, const QJsonObject& jsonObjectB
                             ,QStringList& keyListA,QStringList& keyListB)
{
    getAllKeys(jsonObjectA, keyListA);
    getAllKeys(jsonObjectB, keyListB);
}

void MainWindow::getAllKeys(const QJsonObject& jsonObject, QStringList& keyList)
{
    for (auto it = jsonObject.begin(); it != jsonObject.end(); ++it) {
        QString key = it.key();

        // 将键添加到链表中
        keyList.append(key);

        // 递归处理子对象
        if (it.value().isObject()) {
            getAllKeys(it.value().toObject(), keyList);
        }

        // 递归处理数组中的元素
        if (it.value().isArray()) {
            QJsonArray jsonArray = it.value().toArray();
            for (int i = 0; i < jsonArray.size(); ++i) {
                if (jsonArray.at(i).isObject()) {
                    getAllKeys(jsonArray.at(i).toObject(), keyList);
                }
            }
        }
    }
}


void MainWindow::getSetConfigureFieldsObjects(){

    QString  setfilePath = filePath +"/SetFile/SetConfigureFields.json";
    // 读取配置文件
    QFile configFileObj(setfilePath);
    if (!configFileObj.open(QIODevice::ReadOnly | QIODevice::Text)) {
        qDebug() << "Failed to open config file!";
        return;
    }
    QJsonDocument configDoc = QJsonDocument::fromJson(configFileObj.readAll());
    config = configDoc.object();
    qDebug() <<config.keys();
}

void MainWindow::removeFieldsFromJson( QJsonObject& jsonObject, const QStringList& fieldsToRemove)
{
    for (auto it = jsonObject.begin(); it != jsonObject.end();)
    {
        if (fieldsToRemove.contains(it.key()))
        {
            it = jsonObject.erase(it);
        }
        else
        {
            if (it.value().isObject())
            {
                QJsonObject nestedObject = it.value().toObject();
                removeFieldsFromJson(nestedObject, fieldsToRemove);
                it.value() = nestedObject;
            }
            else if (it.value().isArray())
            {
                QJsonArray jsonArray = it.value().toArray();
                for (int i = 0; i < jsonArray.size(); ++i)
                {
                    if (jsonArray[i].isObject())
                    {
                        QJsonObject nestedObject = jsonArray[i].toObject();
                        removeFieldsFromJson(nestedObject, fieldsToRemove);
                        jsonArray[i] = nestedObject;
                    }
                }
                it.value() = jsonArray;
            }
            ++it;
        }
    }
}


QJsonObject MainWindow::ProcessingObjects(const QJsonObject& jsonObject){

    // 移除字段并重新整合
    QJsonObject newJsonData =jsonObject;
    removeFieldsFromJson(newJsonData, config.keys());
    qDebug() << "Removed JSON: " << newJsonData;

    return newJsonData;
}


//处理对象
QPair<QJsonObject, QJsonObject> MainWindow::getJsonObjects(const QJsonObject& objA, const QJsonObject& objB) {

   QJsonObject newJsonDataA = ProcessingObjects(objA);
   QJsonObject newJsonDataB = ProcessingObjects(objB);

    // 使用 QPair 将两个对象进行组合并返回
    return qMakePair(newJsonDataA, newJsonDataB);
}
void MainWindow::printJsonValueType_B(const QJsonValue& valueA,const QJsonValue& valueB,QString keys,
                                    QTextStream& diffStream,QString pathA,QString pathB,int count) {

    if(valueA != valueB){
        diffStream << "Difference found in field: " << keys <<endl;

        if (valueB.isUndefined()) {
            qDebug() << "ValueA is undefined.";
            diffStream << "Value in file A (" << pathA << "): " << "" << endl;
            diffStream << "Value in file B (" << pathB << "): " << "undefined "<< endl;
            setData(keys, count, "(:undefined )");
            setDifferencequantityData(count, 1);

        } else if (valueB.isNull()) {
            qDebug() << "ValueA is Null.";
            diffStream << "Value in file A (" << pathA << "): " << "" << endl;
            diffStream << "Value in file B (" << pathB << "): " << "Null"<< endl;
            setData(keys, count, "(:Null )");
            setDifferencequantityData(count, 1);

        } else if (valueB.isBool() ) {
            diffStream << "Value in file A (" << pathA << "): " << valueA.toBool() << endl;
            diffStream << "Value in file B (" << pathB << "): " << valueB.toBool() << endl;
            setData(keys, count, "("+QString::number(valueA.isBool())+":"+QString::number(valueB.isBool())+")");
            setDifferencequantityData(count, 1);

        } else if (valueB.isDouble() ) {
            diffStream << "Value in file A (" << pathA << "): " << valueA.toDouble() << endl;
            diffStream << "Value in file B (" << pathB << "): " << valueB.toDouble() << endl;
            setData(keys, count, "("+QString::number(valueA.toInt())+":"+QString::number(valueB.toInt())+")");
            setDifferencequantityData(count, 1);

        } else if (valueB.isString() ) {
            diffStream << "Value in file A (" << pathA << "): " << valueA.toString() << endl;
            diffStream << "Value in file B (" << pathB << "): " << valueB.toString() << endl;
            setData(keys, count, "("+valueA.toString()+":"+valueB.toString()+")");
            setDifferencequantityData(count, 1);

        } else if (valueB.isArray()) {
            diffStream << "ValueA is array." <<endl;
            QJsonArray arrayA = valueA.toArray();
            QJsonArray arrayB = valueB.toArray();
            for (const QJsonValue& arrayValueA : arrayA) {
                for (const QJsonValue& arrayValueB : arrayB) {
                    printJsonValueType(arrayValueA,arrayValueB,keys,diffStream,pathA,pathB,count);
                }
            }
        } else if (valueB.isObject()) {
            diffStream << "ValueA is object."<<endl;
            QJsonObject objectA = valueA.toObject();
            QJsonObject objectB = valueB.toObject();

            QStringList alist = objectA.keys();
            QStringList blist = objectB.keys();

            if(alist.count() >= blist.count()){

                for (const QString& key : objectA.keys()) {
                    printJsonValueType(objectA.value(key),objectB.value(key),key,diffStream,pathA,pathB,count);
                }
            }

            if(alist.count() < blist.count()){
                qDebug() << "objectA.keys() < objectB.keys()";
                for (const QString& key : objectB.keys()) {
                    printJsonValueType_B(objectA.value(key),objectB.value(key),key,diffStream,pathA,pathB,count);
                }
            }
        }
    }
}
void MainWindow::printJsonValueType(const QJsonValue& valueA,const QJsonValue& valueB,QString keys,
                                    QTextStream& diffStream,QString pathA,QString pathB,int count) {

    if(valueA != valueB){
        diffStream << "Difference found in field: " << keys <<endl;

        if (valueA.isUndefined()) {
            qDebug() << "ValueA is undefined.";
            diffStream << "Value in file A (" << pathA << "): " << "undefined" << endl;
            diffStream << "Value in file B (" << pathB << "): " << " "<< endl;
            setData(keys, count, "(undefined: )");
            setDifferencequantityData(count, 1);

        } else if (valueA.isNull()) {
            qDebug() << "ValueA is Null.";
            diffStream << "Value in file A (" << pathA << "): " << "Null" << endl;
            diffStream << "Value in file B (" << pathB << "): " << " "<< endl;
            setData(keys, count, "(Null: )");
            setDifferencequantityData(count, 1);

        } else if (valueA.isBool() ) {
            diffStream << "Value in file A (" << pathA << "): " << valueA.toBool() << endl;
            diffStream << "Value in file B (" << pathB << "): " << valueB.toBool() << endl;
            setData(keys, count, "("+QString::number(valueA.isBool())+":"+QString::number(valueB.isBool())+")");
            setDifferencequantityData(count, 1);

        } else if (valueA.isDouble() ) {
            diffStream << "Value in file A (" << pathA << "): " << valueA.toDouble() << endl;
            diffStream << "Value in file B (" << pathB << "): " << valueB.toDouble() << endl;
            setData(keys, count, "("+QString::number(valueA.toInt())+":"+QString::number(valueB.toInt())+")");
            setDifferencequantityData(count, 1);

        } else if (valueA.isString() ) {
            diffStream << "Value in file A (" << pathA << "): " << valueA.toString() << endl;
            diffStream << "Value in file B (" << pathB << "): " << valueB.toString() << endl;
            setData(keys, count, "("+valueA.toString()+":"+valueB.toString()+")");
            setDifferencequantityData(count, 1);

        } else if (valueA.isArray()) {
            diffStream << "ValueA is array." <<endl;
            QJsonArray arrayA = valueA.toArray();
            QJsonArray arrayB = valueB.toArray();
            for (const QJsonValue& arrayValueA : arrayA) {
                for (const QJsonValue& arrayValueB : arrayB) {
                    printJsonValueType(arrayValueA,arrayValueB,keys,diffStream,pathA,pathB,count);
                }
            }
        } else if (valueA.isObject()) {
            diffStream << "ValueA is object."<<endl;
            QJsonObject objectA = valueA.toObject();
            QJsonObject objectB = valueB.toObject();

            QStringList alist = objectA.keys();
            QStringList blist = objectB.keys();

            if(alist.count() >= blist.count()){

                for (const QString& key : objectA.keys()) {
                    printJsonValueType(objectA.value(key),objectB.value(key),key,diffStream,pathA,pathB,count);
                }
            }

            if(alist.count() < blist.count()){
                qDebug() << "objectA.keys() < objectB.keys()";
                for (const QString& key : objectB.keys()) {
                    printJsonValueType_B(objectA.value(key),objectB.value(key),key,diffStream,pathA,pathB,count);
                }
            }
        }
    }
}


void MainWindow::compareJson(const QJsonObject& objA, const QJsonObject& objB, const QString& pathA, const QString& pathB, QTextStream& diffStream
                             ,int count)
{

    // 调用函数获取两个 QJsonObject 对象,剔除配置文件里的字段
    QPair<QJsonObject, QJsonObject> jsonObjects = getJsonObjects(objA,objB);

    QStringList keys = jsonObjects.first.keys();
    QList<int> firstlist;
    QList<int> secondlist;

    for (const QString& key : keys) {
        if (key == "RootPath" ||
            key == "TotalDefectInfo") {
            continue;
        }

        const QJsonValue& valueA = jsonObjects.first.value(key);
        const QJsonValue& valueB = jsonObjects.second.value(key);

        //若结果数量不一样，返回。
        if(key =="Result" && valueA.toInt() != valueB.toInt()){
            printJsonValueType(valueA,valueB,key,diffStream,pathA,pathB,count);
            return;
        }

        if (valueA != valueB) {
            printJsonValueType(valueA,valueB,key,diffStream,pathA,pathB,count);
        }
    }

//  从返回的 QPair 中分别获取两个对象
    QJsonArray totalDefectInfoA = jsonObjects.first.value("TotalDefectInfo").toArray();
    QJsonArray totalDefectInfoB = jsonObjects.second.value("TotalDefectInfo").toArray();

    firstlist.clear();
    secondlist.clear();

    for (int i = 0; i < totalDefectInfoA.count(); i++) {
        QJsonObject defectInfoA = totalDefectInfoA.at(i).toObject();

        for (int j = 0; j < totalDefectInfoB.count(); j++) {

            QJsonObject defectInfoB = totalDefectInfoB.at(j).toObject();

            if (defectInfoA == defectInfoB) {

                firstlist.append(i);
                secondlist.append(j);
                break;
            }
        }
    }

    if(!firstlist.isEmpty() && !secondlist.isEmpty()){
        //排序大小
//        std::sort(firstlist.begin(), firstlist.end());
//        std::sort(secondlist.begin(), secondlist.end());
        qSort(firstlist);
        qSort(secondlist);

        for(int i = firstlist.count()-1; 0 <= i; --i){

            totalDefectInfoA.removeAt(firstlist.at(i));
            totalDefectInfoB.removeAt(secondlist.at(i));
        }

        firstlist.clear();
        secondlist.clear();

        for (int i = 0; i < totalDefectInfoA.count(); i++) {

            QJsonObject defectInfoA = totalDefectInfoA.at(i).toObject();
            diffStream << "totalDefectInfoA: " << i << ";1"<< endl;

            for (int j = 0; j < totalDefectInfoB.count(); j++) {

                QJsonObject defectInfoB = totalDefectInfoB.at(j).toObject();
                diffStream << "totalDefectInfoB: " << j << endl;


                //若缺陷坐标一样，则对比详情的不一样字段****************//
                if(defectInfoA.value("CenterX").toInt() == defectInfoB.value("CenterX").toInt()
                && defectInfoA.value("CenterY").toInt() == defectInfoB.value("CenterY").toInt()){

                    //记录缺陷序号*************//
                    firstlist.append(i);
                    secondlist.append(j);

                    //对比"AttachInfo"、"RotatedRect"和其它对象数据********//
                    QStringList defectkeysA = defectInfoA.keys();

                    for (const QString& key : defectkeysA) {

                         QJsonValue defevalueA = defectInfoA.value(key);
                         QJsonValue defevalueB = defectInfoB.value(key);
                         printJsonValueType(defevalueA,defevalueB,key,diffStream,pathA,pathB,count);

                    }
                }
            }
        }
    }
    else //当出现对象里除坐标一样，其它不一样时，需要与坐标作为关联
    {

        for (int i = 0; i < totalDefectInfoA.count(); i++) {

            QJsonObject defectInfoA = totalDefectInfoA.at(i).toObject();
            diffStream << "totalDefectInfoA: " << i << ";1"<< endl;
//            QJsonObject CopydefeinfoB;

            for (int j = 0; j < totalDefectInfoB.count(); j++) {

                QJsonObject defectInfoB = totalDefectInfoB.at(j).toObject();
                diffStream << "totalDefectInfoB: " << j << ";1"<< endl;
//                CopydefeinfoB = defectInfoB;

                //若缺陷坐标一样，则对比详情的不一样字段****************//
                if(defectInfoA.value("CenterX").toInt() == defectInfoB.value("CenterX").toInt()
                && defectInfoA.value("CenterY").toInt() == defectInfoB.value("CenterY").toInt()){

                    //记录缺陷序号*************//
                    firstlist.append(i);
                    secondlist.append(j);

                    //对比"AttachInfo"、"RotatedRect"和其它对象数据********//
                    QStringList defectkeysA = defectInfoA.keys();

                    for (const QString& key : defectkeysA) {

                         QJsonValue defevalueA = defectInfoA.value(key);
                         QJsonValue defevalueB = defectInfoB.value(key);
                         printJsonValueType(defevalueA,defevalueB,key,diffStream,pathA,pathB,count);

                    }
                }
            }
        }
    }

    if(!firstlist.isEmpty() && !secondlist.isEmpty()){

        //排序大小
//        std::sort(firstlist.begin(), firstlist.end());
//        std::sort(secondlist.begin(), secondlist.end());
        qSort(firstlist);
        qSort(secondlist);

        for(int i = firstlist.count()-1; 0 <=i; --i){

            totalDefectInfoA.removeAt(firstlist.at(i));
            totalDefectInfoB.removeAt(secondlist.at(i));
        }

        firstlist.clear();
        secondlist.clear();

    }

    //以下就剩下找不到对应坐标时，而匹配的
    for (int i = 0; i < totalDefectInfoA.count(); i++) {

        QJsonObject defectInfoA = totalDefectInfoA.at(i).toObject();
        diffStream << "totalDefectInfoA: " << i << ";2"<<endl;

        for (int j = 0; j < totalDefectInfoB.count(); j++) {

            QJsonObject defectInfoB = totalDefectInfoB.at(j).toObject();
            diffStream << "totalDefectInfoB: " << j << ";2"<< endl;

            //对比"AttachInfo"、"RotatedRect"和其它对象数据********//
            QStringList defectkeysA = defectInfoA.keys();

            for (const QString& key : defectkeysA) {

                 QJsonValue defevalueA = defectInfoA.value(key);
                 QJsonValue defevalueB = defectInfoB.value(key);

                 printJsonValueType(defevalueA,defevalueB,key,diffStream,pathA,pathB,count);

            }
        }
    }
}


void MainWindow::compareJsonFiles(const QString& filePathA, const QString& filePathB, QTextStream& diffStream
                                  ,QString pathfileA,QString pathfileB,int count)
{
    QFile fileA(filePathA);
    QFile fileB(filePathB);

    if (!fileA.open(QIODevice::ReadOnly) ) {
        qDebug() << fileA <<" Failed to open files for reading.";
        return;
    }
    if ( !fileB.open(QIODevice::ReadOnly)) {
        qDebug() << fileB <<" Failed to open files for reading.";
        return;
    }

    QJsonDocument jsonDocA = QJsonDocument::fromJson(fileA.readAll());
    QJsonDocument jsonDocB = QJsonDocument::fromJson(fileB.readAll());

    if (!jsonDocA.isObject() ) {
        qDebug() << "jsonDocA.isObject() " << "Invalid JSON format.";
        return;
    }
    if ( !jsonDocB.isObject()) {
        qDebug() << "jsonDocB.isObject() " << "Invalid JSON format.";
        return;
    }

    QJsonObject jsonObjA = jsonDocA.object();
    QJsonObject jsonObjB = jsonDocB.object();

    setData("测试文件名",count, pathfileA+":"+pathfileB);
    compareJson(jsonObjA, jsonObjB, filePathA, filePathB, diffStream,count);

}

void MainWindow::CheckJsonFiles(const QString& filePathA, const QString& filePathB,bool &MissingAttributeBool)
{
    QFile fileA(filePathA);
    QFile fileB(filePathB);

    if (!fileA.open(QIODevice::ReadOnly) ) {
        qDebug() << fileA <<" Failed to open files for reading. 无法检查缺失属性";
        MissingAttributeBool = true;
        return;
    }
    if ( !fileB.open(QIODevice::ReadOnly)) {
        qDebug() << fileB <<" Failed to open files for reading. 无法检查缺失属性";
        MissingAttributeBool = true;
        return;
    }

    QJsonDocument jsonDocA = QJsonDocument::fromJson(fileA.readAll());
    QJsonDocument jsonDocB = QJsonDocument::fromJson(fileB.readAll());

    if (!jsonDocA.isObject() ) {
        qDebug() << "jsonDocA.isObject() " << "Invalid JSON format. 无法检查缺失属性";
        MissingAttributeBool = true;
        return;
    }
    if ( !jsonDocB.isObject()) {
        qDebug() << "jsonDocB.isObject() " << "Invalid JSON format. 无法检查缺失属性";
        MissingAttributeBool = true;
        return;
    }

    QJsonObject jsonObjA = jsonDocA.object();
    QJsonObject jsonObjB = jsonDocB.object();

    qDebug() << "根据两个路径中的json结构判断属性是否一致";
    //根据两个路径中的json结构判断属性是否一致
    QStringList listA;
    QStringList listB;
    readAllKeys(jsonObjA, jsonObjB,listA,listB);

    QSet<QString> uniqueSetA = listA.toSet();
    QStringList uniqueListA = uniqueSetA.toList();

    QSet<QString> uniqueSetB = listB.toSet();
    QStringList uniqueListB = uniqueSetB.toList();

    QStringList ABdefectstr = findDifferentValues(uniqueListA,uniqueListB);
    qDebug()<< "ABdefectstr: "<<ABdefectstr;

    if(!ABdefectstr.isEmpty()){
        QStringList keydefectstr = findDifferentFields(ABdefectstr,config.keys());
        if(!keydefectstr.isEmpty()){
            MissingAttributeBool = true;
            qDebug() << "缺失属性："<<keydefectstr;
            qDebug() << "请检查两个文件所有属性是否一致，补全剔除属性再进行对比！";
            return;
        }
    }
    MissingAttributeBool = false;
    return;
}


void MainWindow::ObtainFilesWithTheSameStructure(const QStringList& subfoldersA, const QStringList& subfoldersB,
                                                 QString& fileNameA,  QString& fileNameB){

    foreach (const QString& folderA, subfoldersA) {
        QStringList namePartsA = folderA.split("_");

        if (namePartsA.length() > 2 && namePartsA[0] == "NG") {
            QString keyA = namePartsA[0] + "_" + namePartsA[1];

            foreach (const QString& folderB, subfoldersB) {
                QStringList namePartsB = folderB.split("_");

                if (namePartsB.length() > 2 && namePartsB[0] == "NG") {
                    QString keyB = namePartsB[0] + "_" + namePartsB[1];

                    if (keyA == keyB) {
                        fileNameA = keyA +"_" +namePartsA[2];
                        fileNameB = keyB +"_" +namePartsB[2];
                        qDebug() << "fileNameA: "<<fileNameA;
                        qDebug() << "fileNameB: "<<fileNameB;
                        return;
                    }
                }
            }
        }
    }
    return;
}


void MainWindow::CheckAttributes(const QStringList& filePathA, const QStringList& filePathB,const QString& folderPathA, const QString& folderPathB,bool &MissingAttributeBool){

    QString fileNameA;
    QString fileNameB;
    bool boolcheck = false;

    fileNameA.clear();
    fileNameB.clear();
    ObtainFilesWithTheSameStructure(filePathA,filePathB,fileNameA,fileNameB);
    if(!fileNameA.isEmpty() && !fileNameB.isEmpty()){
        boolcheck = true;
    }


    if(boolcheck){
        QString classilyResultA = folderPathA + "/" + fileNameA + jsonpath;
        QString classilyResultB = folderPathB + "/" + fileNameB + jsonpath;
        CheckJsonFiles(classilyResultA,classilyResultB,MissingAttributeBool);
    }else{
        qDebug() << "未找到匹配的NG类型文件，无法检查两个文件属性！";
        MissingAttributeBool = true;
    }

}
void MainWindow::CheckFileNameFormat_1(const QString& folderPathA, bool& BoolFormatFlag)
{
    QDir dirA(folderPathA);

    QStringList subfoldersA = dirA.entryList(QDir::Dirs | QDir::NoDotAndDotDot);

    QStringList ModifyfoldersA;
    // 使用 "_" 分割文件夹名称并提取中间字段作为主键
    foreach (const QString& folder, subfoldersA) {
        QStringList nameParts = folder.split("_");
        if (nameParts.length() > 3) {  // 假设文件夹名字格式为 "prefix_key+prefix_key_suffix"
            ModifyfoldersA.append(folder);
        }
    }


    //修改文件夹名
    foreach (QString folder, ModifyfoldersA) {
        if (folder.contains("+NG_")) {
            QString old_path = dirA.absoluteFilePath(folder);
            QString new_name = folder.replace("+NG_", "");
            QString new_path = dirA.absoluteFilePath(new_name);
            dirA.rename(old_path, new_path);
        } else if (folder.contains("+OK_")) {
            QString old_path = dirA.absoluteFilePath(folder);
            QString new_name = folder.replace("+OK_", "");
            QString new_path = dirA.absoluteFilePath(new_name);
            dirA.rename(old_path, new_path);
        }
        qDebug() << "A: " << "修改完成！";
    }

    BoolFormatFlag = false;
}




void MainWindow::CheckFileNameFormat(const QString& folderPathA, const QString& folderPathB,bool& BoolFormatFlag)
{
    QDir dirA(folderPathA);
    QDir dirB(folderPathB);

    QStringList subfoldersA = dirA.entryList(QDir::Dirs | QDir::NoDotAndDotDot);
    QStringList subfoldersB = dirB.entryList(QDir::Dirs | QDir::NoDotAndDotDot);

    QStringList listfoldersA;
    foreach (QString folder, subfoldersA) {
        if (folder.contains("+NG_")) {
            QString new_name = folder.replace("+NG_", "");
            listfoldersA.append(new_name);
        } else if (folder.contains("+OK_")) {
            QString new_name = folder.replace("+OK_", "");
            listfoldersA.append(new_name);
        }
    }

    QStringList listfoldersB;
    foreach (QString folder, subfoldersB) {
        if (folder.contains("+NG_")) {
            QString new_name = folder.replace("+NG_", "");
            listfoldersB.append(new_name);
        } else if (folder.contains("+OK_")) {
            QString new_name = folder.replace("+OK_", "");
            listfoldersB.append(new_name);
        }
    }

    QStringList sulistfoldersA;
    // 使用 "_" 分割文件夹名称并提取中间字段作为主键
    foreach (const QString& folder, listfoldersA) {
        QStringList nameParts = folder.split("_");
        if (nameParts.length() > 2) {  // 假设文件夹名字格式为 "prefix_key+prefix_key_suffix"
            QString key = nameParts[1];
            sulistfoldersA.append(key);
        }
    }

    QStringList sulistfoldersB;
    // 使用 "_" 分割文件夹名称并提取中间字段作为主键
    foreach (const QString& folder, listfoldersB) {
        QStringList nameParts = folder.split("_");
        if (nameParts.length() > 2) {  // 假设文件夹名字格式为 "prefix_key+prefix_key_suffix"
            QString key = nameParts[1];
            sulistfoldersB.append(key);
        }
    }

    QSet<QString> uniqueSetA = sulistfoldersA.toSet();
    QSet<QString> uniqueSetB = sulistfoldersB.toSet();

    bool ModeDifference = (uniqueSetA == uniqueSetB);
    qDebug() << "ModeDifference = "<<ModeDifference;
    if ( ModeDifference ) {
        qDebug() << "检查文件名及数量ok！";
    } else if(!ModeDifference) {
        qDebug() << "文件夹不匹配数量，请重新组织测试文件后再测！";
        BoolFormatFlag = true;
        return;
    }


    QStringList ModifyfoldersA;
    // 使用 "_" 分割文件夹名称并提取中间字段作为主键
    foreach (const QString& folder, subfoldersA) {
        QStringList nameParts = folder.split("_");
        if (nameParts.length() > 3) {  // 假设文件夹名字格式为 "prefix_key+prefix_key_suffix"
            ModifyfoldersA.append(folder);
        }
    }

    QStringList ModifyfoldersB;
    foreach (const QString& folder, subfoldersB) {
        QStringList nameParts = folder.split("_");
        if (nameParts.length() > 3) {  // 假设文件夹名字格式为 "prefix_key+prefix_key_suffix"
            ModifyfoldersB.append(folder);
        }
    }

    //修改文件夹名
    foreach (QString folder, ModifyfoldersA) {
        if (folder.contains("+NG_")) {
            QString old_path = dirA.absoluteFilePath(folder);
            QString new_name = folder.replace("+NG_", "");
            QString new_path = dirA.absoluteFilePath(new_name);
            dirA.rename(old_path, new_path);
        } else if (folder.contains("+OK_")) {
            QString old_path = dirA.absoluteFilePath(folder);
            QString new_name = folder.replace("+OK_", "");
            QString new_path = dirA.absoluteFilePath(new_name);
            dirA.rename(old_path, new_path);
        }
        qDebug() << "A: " << "修改完成！";
    }

    foreach (QString folder, ModifyfoldersB) {
        if (folder.contains("+NG_")) {
            QString old_path = dirB.absoluteFilePath(folder);
            QString new_name = folder.replace("+NG_", "");
            QString new_path = dirB.absoluteFilePath(new_name);
            dirB.rename(old_path, new_path);
        } else if (folder.contains("+OK_")) {
            QString old_path = dirB.absoluteFilePath(folder);
            QString new_name = folder.replace("+OK_", "");
            QString new_path = dirB.absoluteFilePath(new_name);
            dirB.rename(old_path, new_path);
        }
        qDebug() << "B: " << "修改完成！";
    }

    BoolFormatFlag = false;
}



void MainWindow::processFolder(const QString& folderPathA, const QString& folderPathB, QTextStream& diffStream,bool& MissingAttributeBool)
{
    QDir dirA(folderPathA);
    QDir dirB(folderPathB);

    QStringList subfoldersA = dirA.entryList(QDir::Dirs | QDir::NoDotAndDotDot);
    QStringList subfoldersB = dirB.entryList(QDir::Dirs | QDir::NoDotAndDotDot);

    //排序
    // 使用 "_" 分割文件夹名称并提取中间字段作为主键
    QStringList keysA;
    foreach (const QString& folder, subfoldersA) {
        QStringList nameParts = folder.split("_");
        if (nameParts.length() > 2) {  // 假设文件夹名字格式为 "prefix_key_suffix"
            QString key = nameParts[1];
            keysA.append(key);
        }
    }

    QStringList keysB;
    foreach (const QString& folder, subfoldersB) {
        QStringList nameParts = folder.split("_");
        if (nameParts.length() > 2) {  // 假设文件夹名字格式为 "prefix_key_suffix"
            QString key = nameParts[1];
            keysB.append(key);
        }
    }

    //进行排序
    keysA.sort();
    keysB.sort();

    // 检查B中的所有键是否都存在于A中
    QSet<QString> keySetA = keysA.toSet();
    QSet<QString> keySetB = keysB.toSet();

    bool hasDifference = (keySetA == keySetB);

    qDebug() << "hasDifference = "<<hasDifference;
    if ( hasDifference ) {
        qDebug() << "所有B中的键都存在于A中。";
    } else if(!hasDifference) {
        qDebug() << "文件夹不匹配，请重新组织测试文件后再测。";
        return;
    }



    // 排序
    // 使用 "_" 分割文件夹名称并提取中间字段作为主键
    QMap<QString, QString> folderMapA;
    foreach (const QString& folder, subfoldersA) {
        QStringList nameParts = folder.split("_");
        if (nameParts.length() > 2) {  // 假设文件夹名字格式为 "prefix_key_suffix"
            QString key = nameParts[1];
            folderMapA.insert(key, folder);
        }
    }

    QStringList sortedFoldersA;
    foreach (const QString& key, keysA) {
        if (folderMapA.contains(key)) {
            sortedFoldersA.append(folderMapA.value(key));
        }
    }

    QMap<QString, QString> folderMapB;
    foreach (const QString& folder, subfoldersB) {
        QStringList nameParts = folder.split("_");
        if (nameParts.length() > 2) {  // 假设文件夹名字格式为 "prefix_key_suffix"
            QString key = nameParts[1];
            folderMapB.insert(key, folder);
        }
    }

    QStringList sortedFoldersB;
    foreach (const QString& key, keysB) {
        if (folderMapB.contains(key)) {
            sortedFoldersB.append(folderMapB.value(key));
        }
    }

    //检查属性
    CheckAttributes(sortedFoldersA,sortedFoldersB,folderPathA,folderPathB,MissingAttributeBool);

    if(!MissingAttributeBool){
        // 输出已排序的文件夹列表
        qDebug() << "用keysA的值对subfoldersA的文件夹进行排序：" << sortedFoldersA;
        qDebug() << "用keysB的值对subfoldersB的文件夹进行排序：" << sortedFoldersB;
        for (int i =0;i<sortedFoldersA.count();i++) {
            QString classilyResultA = folderPathA + "/" + sortedFoldersA.at(i) + jsonpath;
            QString classilyResultB = folderPathB + "/" + sortedFoldersB.at(i) + jsonpath;
                compareJsonFiles(classilyResultA, classilyResultB, diffStream,sortedFoldersA.at(i),sortedFoldersB.at(i),2+i);
            }
    }
}

void MainWindow::processFolder_1(const QString& folderPathA,  QTextStream& diffStream)
{
    QDir dirA(folderPathA);

    QStringList subfoldersA = dirA.entryList(QDir::Dirs | QDir::NoDotAndDotDot);

    QString classilyResultA = folderPathA + "/" + subfoldersA.at(0) +jsonpath;

    for (int i =1;i<subfoldersA.count();i++) {
        QString classilyResultB = folderPathA + "/" + subfoldersA.at(i) + jsonpath;

            compareJsonFiles(classilyResultA, classilyResultB, diffStream,subfoldersA.at(0),subfoldersA.at(i),1+i);
        }
}


void MainWindow::openFolder(const QString& folderPath)
{
       QDesktopServices::openUrl(QUrl::fromLocalFile(folderPath));
}


void MainWindow::RecordDataToExcel(QString Path)
{
    try {

    //根据当前时间创建ex表名
//    qDebug()<<Path;
    QDir *DataFile = new QDir;
    bool exist = DataFile->exists("ClassComparedData");
    if(!exist)
    {
        bool isok = DataFile->mkdir("ClassComparedData"); // 新建文件夹
            if(!isok){
                qDebug()<<"创建文件夹ClassComparedData文件夹失败";//提示
                return;
             }
    }
    QDateTime current_date_time =QDateTime::currentDateTime();
    QString current_date =current_date_time.toString("yyyyMMddhhmmss");
    QString absFilePath =Path + "/ClassComparedData/"+current_date+".xlsx";
    EcxelFilePath = absFilePath;
    absFilePath = absFilePath.replace("/", "\\");
//     qDebug()<<absFilePath;
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

     QStringList A_Zlist1;
     QStringList A_Zlist2;
     A_Zlist1<<"A"<<"B"<<"C"<<"D"<<"E"<<"F"<<"G"<<"H"<<"I"<<"J"<<"K"<<"L"<<"M"<<"N"<<"O"<<"P"<<"Q"<<"R"<<"S"<<"T"<<"U"<<"V"<<"W"<<"X"<<"Y"<<"Z";
     A_Zlist2<<"A"<<"B"<<"C"<<"D"<<"E"<<"F"<<"G"<<"H"<<"I"<<"J"<<"K"<<"L"<<"M"<<"N"<<"O"<<"P"<<"Q"<<"R"<<"S"<<"T"<<"U"<<"V"<<"W"<<"X"<<"Y"<<"Z";
     QStringList cellrowlist;
     int size = 0;
     int count = 0;
     int flage = 0;
     QString str;
     if(!ExcelKeylist.isEmpty()){
         for (int i = 0; i < ExcelKeylist.count(); i++) {

             if(count == 26){
                 flage  = flage + 1;
                 count = 0;
                 if(flage > 1){
                     size = size + 1;
                 }
             }

             if(flage >= 1){
                 str = A_Zlist2.at(size) +  A_Zlist1.at(count);
             }
             else {
                 str =  A_Zlist1.at(count);
             }
            cellrowlist<<str;
            count = count + 1;
         }
     }
//      qDebug()<<"设置标题";
     //设置标题
     for(int i = 0; i < cellrowlist.count(); i++){
        QString str = cellrowlist.at(i)+QString::number(1);
        QString Tostr = ExcelKeylist.at(i);
//        qDebug()<<str;
//        qDebug()<<Tostr;

       QAxObject * cell = worksheet->querySubObject("Range(QVariant, QVariant)",str);//获取单元格
        cell->dynamicCall("SetValue(const QVariant&)",QVariant(Tostr));//设置单元格的值
        cell->setProperty("HorizontalAlignment", -4108); //左对齐（xlLeft）：-4131  居中（xlCenter）：-4108  右对齐（xlRight）：-415

     }
      qDebug()<<"开始写入";

     for (int i = 0 ; i < listcountlist.count(); i++) {
         listcount = listcountlist.at(i).toInt();
         rowcount  = rowcountlist.at(i).toInt();
         StrData   = counstrlist.at(i);

         counstr = cellrowlist.at(listcount)+QString::number(rowcount);
//         qDebug()<<"counstr :"<<counstr;
         Colorcell = worksheet->querySubObject("Range(QVariant, QVariant)",counstr);//获取单元格
         var = Colorcell->dynamicCall("Value");
         StrData = var.toString() + StrData;
         Colorcell->dynamicCall("SetValue(const QVariant&)",QVariant(StrData));//设置单元格的值
         Colorcell->setProperty("HorizontalAlignment", -4108); //左对齐（xlLeft）：-4131  居中（xlCenter）：-4108  右对齐（xlRight）：-4152

     }
     workbook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(absFilePath)); //保存到filepath
     workbook->dynamicCall("Close (Boolean)", false);  //关闭文件
     excel->dynamicCall("Quit(void)");  //退出
//     DisPlayTip("输出文件："+absFilePath) ;//提示
     qDebug()<<"输出文件："+absFilePath;
     delete excel;
     excel = nullptr;

    } catch (...){
        qDebug()<<"** 写入Excel出现错误 **";
    return;
    }
}

void MainWindow::setData(int Setlistcount, int Setrowcount, QString SetStrData)
{
    listcountlist.append(QString::number(Setlistcount));
    rowcountlist.append(QString::number(Setrowcount));
    counstrlist.append(SetStrData);
    return;
}

void MainWindow::setData(QString SetKey, int Setrowcount, QString SetStrData)
{


    int listcount =0;
    for(int i=0;i<ExcelKeylist.count();i++){
        if(SetKey == ExcelKeylist.at(i)){
//            qDebug()<<"SetKey: "<<SetKey;
//            qDebug()<<"ExcelKeylist.at(i): "<<ExcelKeylist.at(i);
            listcount = i;
        }
    }

    QString position = map.value(SetKey+";"+QString::number(Setrowcount));
//    qDebug()<<"position: "<<position;

    if(!position.isEmpty() ){
        QString counstr = counstrlist.at(position.toInt()-1);
        counstrlist[position.toInt()-1] = counstr+SetStrData;
//        qDebug()<<"counstr :"<<counstr;

    }else
    {
        listcountlist.append(QString::number(listcount));
        rowcountlist.append(QString::number(Setrowcount));
        counstrlist.append(SetStrData);
//         qDebug()<<"SetKey+QString::number(Setrowcount)******: "<<SetKey+";"+QString::number(Setrowcount);
//         qDebug()<<"QString::number(counstrlist.count())******: "<<QString::number(counstrlist.count());
        map.insert(SetKey+";"+QString::number(Setrowcount),QString::number(counstrlist.count()));
    }

    return;
}

void MainWindow::setDifferencequantityData(int Setrowcount, int SetDifferencequantity)
{
    if(!rowcountlist_1.isEmpty()){
        int count = 0;
        QString number =  rowcountlist_1.last();
        if(number.toInt() ==  Setrowcount){
            QString lastnumber = counstrlist_1.last();
            count = lastnumber.toInt()+SetDifferencequantity;
            counstrlist_1.removeLast();
            counstrlist_1.append(QString::number(count));
        }
        else{
            listcountlist_1.append(QString::number(1));
            rowcountlist_1.append(QString::number(Setrowcount));
            counstrlist_1.append(QString::number(SetDifferencequantity));
        }
    }
    else{
        listcountlist_1.append(QString::number(1));
        rowcountlist_1.append(QString::number(Setrowcount));
        counstrlist_1.append(QString::number(SetDifferencequantity));
    }
    return;

}

void MainWindow::ClearAllDatalist()
{
        listcountlist.clear();
        rowcountlist.clear();
        counstrlist.clear();

        listcountlist_1.clear();
        rowcountlist_1.clear();
        counstrlist_1.clear();

        Set_MainKeylist.clear();
        Set_InformationArrayKeylist.clear();
        Set_AttachInfoKeylist.clear();
        Set_RotatedRectKeylist.clear();

        map.clear();

}

//将所有差异数据添加到ECXEL列表
void MainWindow::DifferencequantityAppendlist()
{
    for (int i =0;i<listcountlist_1.count();i++) {
        listcountlist.append(listcountlist_1.at(i));
        rowcountlist.append(rowcountlist_1.at(i));
        counstrlist.append(counstrlist_1.at(i));
    }
}


void MainWindow::readjsonfile(){

    QString  setfilePath = filePath +"/SetFile/UniversallyResult.json";

     qDebug()<<setfilePath;

    // 读取JSON文件
    QFile file(setfilePath);
    if (!file.open(QIODevice::ReadOnly)) {
        qDebug() <<setfilePath << " Failed to open JSON file";
        return;
    }

    QByteArray jsonData = file.readAll();
    file.close();

    // 解析JSON数据
    QJsonDocument jsonDoc = QJsonDocument::fromJson(jsonData);
    if (jsonDoc.isNull()) {
        qDebug() << "Failed to parse JSON data";
        return;
    }

    valuesList.clear();

    QJsonObject jsonObj = jsonDoc.object();

    // 遍历所有键值对
    for (const QString& topKey : jsonObj.keys()) {
        QJsonValue topLevelValue = jsonObj.value(topKey);

        if (topLevelValue.isObject()) {
            QJsonObject secondLevelObj = topLevelValue.toObject();
            for (const QString& secondKey : secondLevelObj.keys()) {
                QJsonValue secondLevelValue = secondLevelObj.value(secondKey);
                if (secondLevelValue.isString()) {
                    valuesList.append(secondLevelValue.toString());
                }
            }
        }
    }
}


void MainWindow::setExcelTebel()
{
    ExcelKeylist.clear();//清空
    ExcelKeylist<<"测试文件名"<<"差异数量";

    for(int i = 0 ; i < valuesList.count();i++){
        ExcelKeylist.push_back(valuesList.at(i));
    }

    // 输出链表中的所有值
    for (const QString& value : ExcelKeylist) {
        qDebug() << value;
    }

}

//选择单个log路径
void MainWindow::on_pushButton_clicked()
{
    // 读取上一次选择的路径
    QString lastFolderPath = settings_only.value("LastFolderPath").toString();

    // 打开文件夹对话框，设置初始路径为上一次选择的路径
    QString path = QFileDialog::getExistingDirectory(this, "选择文件夹", lastFolderPath);

    ui->lineEdit->setText(path);
    folderPath = path;

    checkResultLJson(path);

    if (!folderPath.isEmpty())
    {
        // 处理选择的文件夹路径

        // 保存当前选择的路径作为上一次选择的路径
        settings_only.setValue("LastFolderPath", folderPath);
    }

}
//开始对比单个log
void MainWindow::on_pushButton_4_clicked()
{
    if(folderPath.isEmpty())
    {
        qDebug() << "请选择路径";
        return;
    }

    ClearAllDatalist();

//  先读取配置文件中的字段
    readjsonfile();

    setExcelTebel();

    QDateTime current_date_time =QDateTime::currentDateTime();
    QString current_date =current_date_time.toString("yyyyMMddhhmmss");
    QDir *DataFile = new QDir;
    bool exist = DataFile->exists("ComparedDataTxt");
    if(!exist)
    {
        bool isok = DataFile->mkdir("ComparedDataTxt"); // 新建文件夹
            if(!isok){
               qDebug()<<"创建文件夹ComparedDataTxt文件夹失败";//提示
                return;
             }
    }

    QFile outputFile(filePath+ "/ComparedDataTxt/"+current_date+"data.txt");
    if (!outputFile.open(QIODevice::WriteOnly | QIODevice::Text)) {
        qDebug() << "outputFile Failed to open output file.";
    }

    QTextStream diffStream(&outputFile);

    if(ModecomboBox == 2){

        bool falsefolder = false;

        CheckFileNameFormat_1(folderPath,falsefolder);
    }

    processFolder_1(folderPath, diffStream);

    DifferencequantityAppendlist();

    RecordDataToExcel(filePath);

    outputFile.close();

}



//选择log1路径
void MainWindow::on_pushButton_2_clicked()
{

    QString lastFolderPath = settings_1.value("LastFolderPath").toString();

    QString path = QFileDialog::getExistingDirectory(this, "选择文件夹", lastFolderPath);
    ui->lineEdit_2->setText(path);
    folderPathA = path;

    checkResultLJson(path);

    if (!folderPathA.isEmpty())
    {
        // 处理选择的文件夹路径

        // 保存当前选择的路径作为上一次选择的路径
        settings_1.setValue("LastFolderPath", folderPathA);
    }
}

//选择log2路径
void MainWindow::on_pushButton_3_clicked()
{
    QString lastFolderPath = settings_2.value("LastFolderPath").toString();

    QString path = QFileDialog::getExistingDirectory(this, "选择文件夹", lastFolderPath);
    ui->lineEdit_3->setText(path);
    folderPathB = path;

    checkResultLJson(path);

    if (!folderPathB.isEmpty())
    {
        // 处理选择的文件夹路径

        // 保存当前选择的路径作为上一次选择的路径
        settings_2.setValue("LastFolderPath", folderPathB);
    }
}

//开始对比两个log
void MainWindow::on_pushButton_5_clicked()
{

    if(folderPathA.isEmpty() ||folderPathB.isEmpty())
    {
        qDebug() << "请选择路径";
        return;
    }

    ClearAllDatalist();

//  先读取配置文件中的字段
    readjsonfile();

    setExcelTebel();

//    qDebug()<<"filePath "<<filePath<<"\n";
    QDateTime current_date_time =QDateTime::currentDateTime();
    QString current_date =current_date_time.toString("yyyyMMddhhmmss");
    QDir *DataFile = new QDir;
    bool exist = DataFile->exists("ComparedDataTxt");
    if(!exist)
    {
        bool isok = DataFile->mkdir("ComparedDataTxt"); // 新建文件夹
            if(!isok){
               qDebug()<<"创建文件夹ComparedDataTxt文件夹失败";//提示
                return;
             }
    }

    QFile outputFile(filePath+ "/ComparedDataTxt/"+current_date+"data.txt");


    if (!outputFile.open(QIODevice::WriteOnly | QIODevice::Text)) {
        qDebug() << "outputFile Failed to open output file.";
    }

    QTextStream diffStream(&outputFile);

    if(ModecomboBox == 2){

        bool BoolFormatFlag = false;

        CheckFileNameFormat(folderPathA, folderPathB,BoolFormatFlag);

        if(!BoolFormatFlag){

            bool Attributebool = false;

            processFolder(folderPathA, folderPathB, diffStream,Attributebool);

            if(!Attributebool){

                DifferencequantityAppendlist();

                RecordDataToExcel(filePath);
            }

        }

        outputFile.close();

    }else {

        bool Attributebool = false;

        processFolder(folderPathA, folderPathB, diffStream,Attributebool);

        if(!Attributebool){

            DifferencequantityAppendlist();

            RecordDataToExcel(filePath);
        }

        outputFile.close();
    }
}

void MainWindow::on_comboBox_currentIndexChanged(int index)
{
    comboBoxIndex = ui->comboBox->itemText(index);

    if("中间结果左" == comboBoxIndex){
        jsonpath = "/resultL.json";
    }else if("中间结果右" == comboBoxIndex){
        jsonpath = "/resultR.json";
    }else if("分类结果左" == comboBoxIndex){
        jsonpath = "/ClassilyResultL.json";
    }else if("分类结果右" == comboBoxIndex){
        jsonpath = "/ClassilyResultR.json";
    }else if("中间结果" == comboBoxIndex){
        jsonpath = "/result.json";
    }else if("分类结果" == comboBoxIndex){
        jsonpath = "/ClassilyResult.json";
    }
//    jsonpath = comboBoxIndex;
    qDebug() << jsonpath;
}

void MainWindow::on_comboBox_2_currentIndexChanged(int index)
{
    comboBoxIndex_2 = ui->comboBox_2->itemText(index);
//    qDebug() << "comboBoxIndex_2:"<<comboBoxIndex_2;
    if("正常log" == comboBoxIndex_2){
        ModecomboBox = 1;
        // 当comboBox_2选择"正常log"时，设置comboBox的选项为"/result.json"和"/ClassilyResult.json"
        ui->comboBox->clear();
        ui->comboBox->addItem("中间结果");
        ui->comboBox->addItem("分类结果");
    }else if("左右log" == comboBoxIndex_2){
        ModecomboBox = 2;
        // 当comboBox_2选择"左右log"时，设置comboBox的选项为"/resultL.json"、"/resultR.json"、"/ClassilyResultL.json"和"/ClassilyResultR.json"
        ui->comboBox->clear();
        ui->comboBox->addItem("中间结果左");
        ui->comboBox->addItem("中间结果右");
        ui->comboBox->addItem("分类结果左");
        ui->comboBox->addItem("分类结果右");
    }
    qDebug() << ModecomboBox;
}

void MainWindow::handleCleanup(QAxObject* workbook) {

    if (workbook != nullptr && !workbook->isNull()) {
        workbook->dynamicCall("Close()");
    }

    if (lookexcel != nullptr && !lookexcel->isNull()) {
        lookexcel->dynamicCall("Quit()");
    }
    lookexcel = nullptr;
}


void MainWindow::on_pushButton_6_clicked()
{
    if(!EcxelFilePath.isEmpty()){

        lookexcel = new QAxObject(this);//建立excel操作对象
        lookexcel->setControl("Excel.Application");//连接Excel控件
        lookexcel->dynamicCall("SetVisible (bool Visible)","true");//设置为显示窗体
        lookexcel->setProperty("DisplayAlerts", false);//不显示任何警告信息，如关闭时的是否保存提示

        QAxObject *workbooks = lookexcel->querySubObject("Workbooks");
        QAxObject *workbook = workbooks->querySubObject("Open(const QString&)", EcxelFilePath);
        // 检查是否成功打开 Excel 文件
        if (workbook == nullptr || workbook->isNull()) {
            qDebug() << "无法打开 Excel 文件";
            return;
        }

        // 可以在这里进行其他操作，如读取数据等

        QObject::connect(qApp, &QCoreApplication::aboutToQuit, [this, workbook]() {
            handleCleanup(workbook);
        });
    }

}


void MainWindow::checkResultLJson(QString path)
{
    QDir dir(path);
    if (!dir.exists()) {
        qWarning() << "该文件夹" << path << "不存在";
        return;
    }

    QStringList subdirs = dir.entryList(QDir::Dirs | QDir::NoDotAndDotDot, QDir::Name);
    foreach (const QString &subdir, subdirs) {
        QString subPath = QString("%1/%2").arg(path).arg(subdir);

        // 查找 resultL.json 文件
        QFile file(QString("%1"+jsonpath).arg(subPath));
        if (!file.exists()) {
            qWarning() << subPath << "不存在 "+jsonpath+" 文件";
        }
    }
}
