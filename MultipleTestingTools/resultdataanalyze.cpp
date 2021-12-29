#include "resultdataanalyze.h"
#include "ui_resultdataanalyze.h"

ResultDataAnalyze::ResultDataAnalyze(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::ResultDataAnalyze)
{
    ui->setupUi(this);
    this->setFixedSize(818,593);
    filePath = QApplication::applicationDirPath();

    recrdlog = new RecordLog;//动态分配空间，不能指定父对象
    Sub_Thread = new QThread(this);//创建子线程
    recrdlog->moveToThread(Sub_Thread);//把自定义的线程加入子线程
    connect(this,SIGNAL(Startsignal()),recrdlog,SLOT(ThreadLog()));//日志线程
    connect(Sub_Thread, SIGNAL(finished()), recrdlog, SLOT(deleteLater())); //释放线程

    Sub_Thread->start(); //启动子线程
    emit Startsignal();
    SetFileTabel();//讀取配置文件
}

ResultDataAnalyze::~ResultDataAnalyze()
{
    Sub_Thread->quit();
    Sub_Thread->wait();//线程退出--不可delete线程对象
    delete ui;
}


//读取log1
void ResultDataAnalyze::on_PathButton_clicked()
{
    initlog();//初始化日志

    if(Set_MainKeylist.isEmpty() || Set_InformationArrayKeylist.isEmpty()
       || Set_AttachInfoKeylist.isEmpty() || Set_RotatedRectKeylist.isEmpty())
    {
        DisPlayTip("配置文件内容有缺，请按要求配置好文件!");
        return;
    }

    //判断是否选中LOG文件
    QString path = QFileDialog::getExistingDirectory(this, tr("选择文件夹"),"D:\\",QFileDialog::ShowDirsOnly);
    ui->PathlineEdit->setText(path);
    dir = new QDir(path);

    if("LOG"!=dir->dirName() )
    {
        DisPlayTip("请选择正确的路径");
        return;
    }

    ui->PathlineEdit_2->clear();
    FilePathlistlog.clear();

    dircount_1 = 0;
    FilePathlist.clear();
    Filenamelist.clear();

    Q_mainValue_list_1.clear();
    Q_mainKey_list_1.clear();
    Q_informationArrayValue_vectorlist_1.clear();
    Q_AttachInfoValue_vectorlist_1.clear();
    Q_RotatedRectValue_vectorlist_1.clear();
    Q_informationArrayKey_vectorlist_1.clear();
    Q_AttachInfoKey_vectorlist_1.clear();
    Q_RotatedRectKey_vectorlist_1.clear();

    ui->InforEdit_1->clear();


     QString DirPath =dir->QDir::absolutePath();
     QStringList l1=dir->entryList(QDir::Dirs,QDir::Name); //获取当前文件夹的数量
     sort(l1);//排序文件**需要文件分隔后为整数
     for (int i = 0;i < nl1.size(); i++)
     {
       QString fliename = Findfliename(nl1.at(i),l1);
       qDebug()<<"fliename"<<fliename;
       FilePathlist.append(DirPath+"/"+fliename+"/"+SetFileName);//拼接路径
       Filenamelist.append(fliename);
       ui->InforEdit_1->append(fliename);
       dircount_1 = dircount_1 + 1;
     }
     DisPlayTip("开始读取第一个log数据");
     bool bfalse =  getResultdatas(FilePathlist,Q_mainValue_list_1,Q_mainKey_list_1,Q_informationArrayValue_vectorlist_1,
                    Q_AttachInfoValue_vectorlist_1,Q_RotatedRectValue_vectorlist_1,Q_informationArrayKey_vectorlist_1,
                    Q_AttachInfoKey_vectorlist_1,Q_RotatedRectKey_vectorlist_1);//获取第一个log的每个文件夹Result
     if(bfalse){
         setExcelTebel(Q_mainKey_list_1,Q_informationArrayKey_vectorlist_1,Q_AttachInfoKey_vectorlist_1,Q_RotatedRectKey_vectorlist_1);
         DisPlayTip("读取完成");
         }
     else{
         FilePathlist.clear();
         DisPlayTip("读取失败");
     }
}
QString ResultDataAnalyze::Findfliename(int nlist, QStringList list)
{
    QStringList l1;
    QString fliename;

    for (int i = 0; i < list.size(); i++) {
        if("." == list.at(i) || ".." == list.at(i))//过滤.和..隐藏文件
            continue;
        l1.clear();
        fliename = list.at(i);
        l1 = fliename.split("_");
        if(nlist == l1.at(1).toInt())
            return fliename;
    }

}
void ResultDataAnalyze::sort(QStringList list)
{
    QStringList l1;
    QString str;
    nl1.clear();
    for(int i = 0; i < list.size(); i++){
        if("." == list.at(i) || ".." == list.at(i))//过滤.和..隐藏文件
            continue;
        l1.clear();
        str = list.at(i);
        l1 = str.split("_");
        nl1.append(l1.at(1).toInt());
    }
    qSort(nl1.begin(),nl1.end());
    qDebug()<<"qsort"<<nl1;

}

//读取log2
void ResultDataAnalyze::on_PathButton_1_clicked()
{
    initlog();//初始化日志

    if(Set_MainKeylist.isEmpty() || Set_InformationArrayKeylist.isEmpty()
       || Set_AttachInfoKeylist.isEmpty() || Set_RotatedRectKeylist.isEmpty())
    {
        DisPlayTip("配置文件内容有缺，请按要求配置好文件!");
        return;
    }
    //判断是否选中LOG文件
    QString path = QFileDialog::getExistingDirectory(this, tr("选择文件夹"),"D:\\",QFileDialog::ShowDirsOnly);
    ui->PathlineEdit_1->setText(path);
    dir_2 = new QDir(path);
    if("LOG"!=dir_2->dirName())
    {
        DisPlayTip("请选择正确的路径");
        return;
    }

    ui->PathlineEdit_2->clear();
    FilePathlistlog.clear();

    dircount_2 = 0;
    FilePathlist_2.clear();
    Filenamelist_2.clear();

    Q_mainValue_list_2.clear();
    Q_mainKey_list_2.clear();
    Q_informationArrayValue_vectorlist_2.clear();
    Q_AttachInfoValue_vectorlist_2.clear();
    Q_RotatedRectValue_vectorlist_2.clear();
    Q_informationArrayKey_vectorlist_2.clear();
    Q_AttachInfoKey_vectorlist_2.clear();
    Q_RotatedRectKey_vectorlist_2.clear();

    ui->InforEdit_2->clear();

    QString DirPath =dir_2->QDir::absolutePath();
    QStringList l1=dir_2->entryList(QDir::Dirs,QDir::Name); //获取当前文件夹的数量
    sort(l1);//排序文件**需要文件分隔后为整数
     for (int i = 0;i < nl1.size(); i++)
    {
       QString fliename = Findfliename(nl1.at(i),l1);
       qDebug()<<"fliename"<<fliename;
       FilePathlist_2.append(DirPath+"/"+fliename+"/"+SetFileName);//拼接路径
       Filenamelist_2.append(fliename);
       ui->InforEdit_2->append(fliename);
       dircount_2 = dircount_2 + 1;

     }

     DisPlayTip("开始读取第二个log数据");
     bool bfalse =  getResultdatas(FilePathlist_2,Q_mainValue_list_2,Q_mainKey_list_2,Q_informationArrayValue_vectorlist_2,Q_AttachInfoValue_vectorlist_2,
                    Q_RotatedRectValue_vectorlist_2,Q_informationArrayKey_vectorlist_2,
                    Q_AttachInfoKey_vectorlist_2,Q_RotatedRectKey_vectorlist_2);//获取第二个log的每个文件夹Result
     if(bfalse)
         DisPlayTip("读取完成");
     else{
         FilePathlist_2.clear();
         DisPlayTip("读取失败");
     }
}

//对比数据
void ResultDataAnalyze::on_StartButton_clicked()
{
    if(FilePathlist_2.isEmpty() || FilePathlist.isEmpty())
    {
        DisPlayTip("请选择路径");
        return;
    }
    QString tip ;
    QString ReCt ;
    QString filename1;
    QString filename2;
    if(dircount_1 != dircount_2)//对比文件数量
    {
        DisPlayTip("文件夹数量 不匹配");
        recrdlog->write();
        return;
    }


    QStringList list1;
    QStringList list2;
    QString strname1 ;
    QString strname2 ;

    QStringList mainKeylist1   ;
    QStringList mainValuelist1 ;
    QStringList mainValuelist2 ;

    QList<QStringList> inforarrayKey ;
    QList<QStringList> inforarray1   ;
    QList<QStringList> inforarray2   ;
    QStringList  inforarrayKeylist   ;
    QStringList  inforarrayValuelist1;
    QStringList  inforarrayValuelist2;

    QList<QStringList> AttachInfoKey     ;
    QList<QStringList> AttachInfoValue1  ;
    QList<QStringList> AttachInfoValue2  ;
    QStringList AttachInfoKeylist;
    QStringList AttachInfo1;
    QStringList AttachInfo2;

    QList<QStringList> RotatedRectKey   ;
    QList<QStringList> RotatedRectValue1 ;
    QList<QStringList> RotatedRectValue2 ;
    QStringList RotatedRectKeylist;
    QStringList RotatedRect1   ;
    QStringList RotatedRect2  ;

    listcountlist.clear();
    rowcountlist.clear();
    counstrlist.clear();

    try {
        int row = 1;
        int defectcount;//差异数量计数
        int nRestul = 0;
        int H = 0;
        for(int i = 0 ; i < dircount_1; i++)
        {
            tip = QString::number(i);
            row = row + 1;
            defectcount = 0;
            nRestul  = 0 ;
            list1.clear();
            list2.clear();
            filename1 = Filenamelist.at(i);
            filename2 = Filenamelist_2.at(i);
            list1     = filename1.split("_");
            list2     = filename2.split("_");
            filename1 = list1.at(1);
            filename2 = list2.at(1);
            recrdlog->strlog += Filenamelist.at(i) + " ** " + Filenamelist_2.at(i) + "\n";
            qDebug()<<filename1;
            qDebug()<<filename2;
            setData(0,row, Filenamelist.at(i) + ";" + Filenamelist_2.at(i));
            if(filename1 != filename2)
            {
                DisPlayTip(filename1+" ** "+ tip +" "+ filename2 +" 不匹配 ");
//                return;
            }

            mainKeylist1.clear();
            mainValuelist1.clear();
            mainValuelist2.clear();
            mainKeylist1   = Q_mainKey_list_1.at(i);
            mainValuelist1 = Q_mainValue_list_1.at(i);
            mainValuelist2 = Q_mainValue_list_2.at(i);
            H = 2;
            if(mainValuelist1.count() == mainValuelist2.count()){
               for(int x = 0;  x < mainValuelist1.count(); x++){
                   strname1 = mainValuelist1.at(x);
                   strname2 = mainValuelist2.at(x);
                   if(strname1 != strname2){

                       DisPlayTip(filename1+" ** mainKey "+ QString::number(H + x) +" "+ mainKeylist1.at(x) +" 不匹配 ");
                       setData(H + x,row, strname1 + ";" + strname2);
                       defectcount++;

                       if("Result" == mainKeylist1.at(x)){
                            nRestul = 1;
                       }
                   }
               }
               if(1 == nRestul)
                   DisPlayTip("Result 数量不一致");
                   setData(1,row, QString::number(defectcount));
                   continue;
            }

            inforarrayKey.clear();
            inforarray1.clear();
            inforarray2.clear();
            inforarrayKey = Q_informationArrayKey_vectorlist_1.at(i);
            inforarray1   = Q_informationArrayValue_vectorlist_1.at(i);
            inforarray2   = Q_informationArrayValue_vectorlist_2.at(i);
            H = 2 + mainKeylist1.count();
            if(inforarray1.count() == inforarray2.count()){

                for(int D = 0 ; D < inforarray1.count(); D++){
                    ReCt = QString::number(D);

                    inforarrayKeylist.clear();
                    inforarrayValuelist1.clear();
                    inforarrayValuelist2.clear();
                    inforarrayKeylist    = inforarrayKey.at(D);
                    inforarrayValuelist1 = inforarray1.at(D);
                    inforarrayValuelist2 = inforarray2.at(D);

                        if(inforarrayValuelist1.count() == inforarrayValuelist2.count()){
                               for(int x = 0; x < inforarrayValuelist1.count(); x++){
                                   strname1 = inforarrayValuelist1.at(x);
                                   strname2 = inforarrayValuelist2.at(x);

                                       if(strname1 != strname2){
                                           DisPlayTip(filename1+" ** inforarrayKey "+ QString::number(H + x) +" "+ inforarrayKeylist.at(x) +" 不匹配 ");
                                           setData(H + x,row, ReCt+" -> "+strname1 + ";" + strname2+"\n");
                                           defectcount++;
                                        }
                                }
                          }
                    }
             }
//             else {
//                //報錯缺陷信息數目不一致
//                DisPlayTip("inforarray 差异数量不一致");
//                setData(1,row, QString::number(defectcount));
//                continue;
//             }
            AttachInfoKey.clear();
            AttachInfoValue1.clear();
            AttachInfoValue2.clear();
            AttachInfoKey      = Q_AttachInfoKey_vectorlist_1.at(i);
            AttachInfoValue1   = Q_AttachInfoValue_vectorlist_1.at(i);
            AttachInfoValue2   = Q_AttachInfoValue_vectorlist_2.at(i);
             H = 2 + mainKeylist1.count() + inforarrayKeylist.count();
            if(AttachInfoValue1.count() == AttachInfoValue2.count()){
                for(int D = 0 ; D < AttachInfoValue1.count(); D++){
                    ReCt = QString::number(D);

                    AttachInfoKeylist.clear();
                    AttachInfo1.clear();
                    AttachInfo2.clear();
                    AttachInfoKeylist = AttachInfoKey.at(D);
                    AttachInfo1       = AttachInfoValue1.at(D);
                    AttachInfo2       = AttachInfoValue2.at(D);

                    if(AttachInfo1.count() == AttachInfo2.count()){
                        for(int x = 0; x < AttachInfo1.count(); x++){
                            strname1 = AttachInfo1.at(x);
                            strname2 = AttachInfo2.at(x);

                            if(strname1 != strname2){

                                DisPlayTip(filename1+" ** AttachInfoKey "+  QString::number(H + x) +" "+ AttachInfoKeylist.at(x) +" 不匹配 ");
                                setData(H + x,row, ReCt+" -> " + strname1 + ";" + strname2+"\n");
                                defectcount++;
                            }
                       }
                  }
               }
            }
//            else {
//                //報錯缺陷信息數目不一致
//                DisPlayTip("AttachInfo 差异数量不一致");
//                setData(1,row, QString::number(defectcount));
//                continue;
//            }

            RotatedRectKey.clear();
            RotatedRectValue1.clear();
            RotatedRectValue2.clear();
            RotatedRectKey      = Q_RotatedRectKey_vectorlist_1.at(i);
            RotatedRectValue1   = Q_RotatedRectValue_vectorlist_1.at(i);
            RotatedRectValue2   = Q_RotatedRectValue_vectorlist_2.at(i);
             H = 2 + mainKeylist1.count() + inforarrayKeylist.count() + AttachInfoKeylist.count();
            if(RotatedRectValue1.count() == RotatedRectValue2.count()){

                   for(int D = 0 ; D < RotatedRectValue1.count(); D++){
                        ReCt = QString::number(D);
                        RotatedRectKeylist.clear();
                        RotatedRect1.clear();
                        RotatedRect2.clear();
                        RotatedRectKeylist = RotatedRectKey.at(D);
                        RotatedRect1       = RotatedRectValue1.at(D);
                        RotatedRect2       = RotatedRectValue2.at(D);
                     if(RotatedRect1.count() == RotatedRect2.count()){
                         for(int x = 0; x < RotatedRect1.count(); x++){

                             strname1 = RotatedRect1.at(x);
                             strname2 = RotatedRect2.at(x);
                             if(strname1 != strname2){
                                 DisPlayTip(filename1+" ** RotatedRectKey "+ QString::number(H + x) +" "+ RotatedRectKeylist.at(x) +" 不匹配  ");
                                 setData(H + x,row, ReCt+" -> " + strname1 + ";" + strname2+"\n");
                                 defectcount++;
                             }
                         }
                     }
                }
            }
//            else {
//                //報錯缺陷信息數目不一致
//                DisPlayTip("RotatedRect 差异数量不一致");
//                setData(1,row, QString::number(defectcount));
//                continue;
//            }
            setData(1,row, QString::number(defectcount));
        }

    } catch (... ){
        DisPlayTip("**对比出现错误**");
        recrdlog->write();
        return;
    }
    DisPlayTip("对比完成");

    RecordDataToExcel(filePath);
    recrdlog->write();
}



//选择單個log
void ResultDataAnalyze::on_pushButton_clicked()
{
    initlog();//初始化日志

    if(Set_MainKeylist.isEmpty() || Set_InformationArrayKeylist.isEmpty()
       || Set_AttachInfoKeylist.isEmpty() || Set_RotatedRectKeylist.isEmpty())
    {
        DisPlayTip("配置文件内容有缺，请按要求配置好文件!");
        return;
    }
    //分析路径 读取文件
    //判断是否选中LOG文件
    QString path = QFileDialog::getExistingDirectory(this, tr("选择文件夹"),"D:\\",QFileDialog::ShowDirsOnly);
    ui->PathlineEdit_2->setText(path);
    dirlog = new QDir(path);

    if("LOG"!=dirlog->dirName() )
    {
        DisPlayTip("请选择正确的路径");
        return;
    }

    ui->PathlineEdit->clear();
    ui->PathlineEdit_1->clear();
    ui->InforEdit_1->clear();
    ui->InforEdit_2->clear();

    Q_mainValue_list.clear() ;
    Q_mainKey_list.clear() ;
    Q_informationArrayValue_vectorlist.clear();
    Q_AttachInfoValue_vectorlist.clear();
    Q_RotatedRectValue_vectorlist.clear();

    Q_informationArrayKey_vectorlist.clear();
    Q_AttachInfoKey_vectorlist.clear();
    Q_RotatedRectKey_vectorlist.clear();

    ui->inforEdit->clear();

    FilePathlist.clear();
    FilePathlist_2.clear();

    FilePathlistlog.clear();
    Filenamelistlog.clear();

      QString DirPath =dirlog->QDir::absolutePath();
      QStringList l1=dirlog->entryList(QDir::Dirs,QDir::Name); //获取当前文件夹的数量
      sort(l1);//排序文件**需要文件分隔后为整数
      for (int i = 0;i < l1.size(); i++)
     {
       if("." == l1.at(i) || ".." == l1.at(i))//过滤.和..隐藏文件
           continue;
       Filenamelistlog.append(l1.at(i));
     }
      for(int i = 0 ; i < nl1.count(); i++){      //去掉重文件名
        for(int j = i + 1; j < nl1.count(); j++){
            if(nl1.at(i) == nl1.at(j)){
                nl1.removeAt(j);
                j--;
            }
        }
      }
      QString filename;
      QStringList filenamelist;
     qDebug()<<nl1;
     NameCountlist.clear();
     Namelist.clear();
     int count;
     for(int i = 0; i < nl1.count(); i++){
         count = 0;
         for(int j = 0 ; j < Filenamelistlog.count(); j++){  //算出相同文件名有多少
             filename = Filenamelistlog.at(j);
             filenamelist.clear();
             filenamelist = filename.split("_");
             if(filenamelist.at(1).toInt() == nl1.at(i)) {
                count = count + 1;
             }
         }
         NameCountlist.append(QString::number(count));
     }


     qDebug()<<NameCountlist.count();
     int namecount = 0;
     for(int i = 0 ; i < NameCountlist.count(); i++)
     {
         namecount = NameCountlist.at(i).toInt();
         if(1 == namecount)
             continue;
         else{
             for(int j = 0 ; j < Filenamelistlog.count(); j++){  //根据文件名读取log数据
                 filename = Filenamelistlog.at(j);
                 filenamelist.clear();
                 filenamelist = filename.split("_");
                     if(filenamelist.at(1).toInt() == nl1.at(i)) {
                        FilePathlistlog.append(DirPath+"/"+Filenamelistlog.at(j)+"/"+SetFileName);//拼接路径
                        ui->InforEdit_1->append(Filenamelistlog.at(j));
                        Namelist.append(Filenamelistlog.at(j));
                     }
                 }
         }
     }

     if(FilePathlistlog.isEmpty()){
         DisPlayTip(path+"不存在对比数据，请选择正确的对比路径");
         return;
     }
     DisPlayTip("开始读取数据");
     bool bfalse = getResultdatas(FilePathlistlog,
                    Q_mainValue_list,Q_mainKey_list,
                    Q_informationArrayValue_vectorlist,Q_AttachInfoValue_vectorlist,Q_RotatedRectValue_vectorlist
                    ,Q_informationArrayKey_vectorlist,Q_AttachInfoKey_vectorlist,Q_RotatedRectKey_vectorlist);

     if(bfalse){
     setExcelTebel(Q_mainKey_list,Q_informationArrayKey_vectorlist,Q_AttachInfoKey_vectorlist,Q_RotatedRectKey_vectorlist);
     DisPlayTip("读取完成");
     }else {
      DisPlayTip("读取失败");
      FilePathlistlog.clear();
     }
    return;
}


//对比单个log数据
void ResultDataAnalyze::on_StartButton_2_clicked()
{
    if(FilePathlistlog.isEmpty())
    {
        DisPlayTip("请选择单个log路径");
        return;
    }

    ResultDataCompared();//对比

    DisPlayTip("对比完成");

    RecordDataToExcel(filePath);//写入excel
    recrdlog->write();
   return;
}

void ResultDataAnalyze::setData(int Setlistcount, int Setrowcount, QString SetStrData)
{
    listcountlist.append(QString::number(Setlistcount));
    rowcountlist.append(QString::number(Setrowcount));
    counstrlist.append(SetStrData);
    return;

}

void ResultDataAnalyze::DisPlayTip(QString strTipstr)
{
    QString logstr ;
    QDateTime da_time;
    QString time_str = da_time.currentDateTime().toString("yyyy-MM-dd HH:mm:ss");
    logstr = "["+time_str+"]" +strTipstr +"\n";;
    ui->inforEdit->append(logstr);
    recrdlog->strlog  += logstr;
    return;
}

void ResultDataAnalyze::closeEvent(QCloseEvent *event)
{

    recrdlog->quit();
    emit CloseResultDataAnlyze();
    qDebug()<<"send CloseResultDataAnlyze ok";

}

//配置文件
void ResultDataAnalyze::on_pushButton_2_clicked()
{
    SetFileTabel();//讀取配置文件
}

bool ResultDataAnalyze::getResultdatas(QStringList &filePathlistlog,QList<QStringList> &mainValue_list,QList<QStringList> &mainKey_list,
 QVector<QList<QStringList>> &informationArrayValue_vectorlist,QVector<QList<QStringList>> &AttachInfoValue_vectorlist, QVector<QList<QStringList>> &RotatedRectValue_vectorlist,
QVector<QList<QStringList>> &informationArrayKey_vectorlist,QVector<QList<QStringList>> &AttachInfoKey_vectorlist,QVector<QList<QStringList>> &RotatedRectKey_vectorlist)
{

    QStringList MainValuelist;
    QStringList MainKeylist;

    QStringList InformationArrayValuelist;
    QStringList AttachInfoValuelist;
    QStringList RotatedRectValuelist;


    QStringList InformationArrayKeylist;
    QStringList AttachInfoKeylist;
    QStringList RotatedRectKeylist;

    //记录数组内容
    QList<QStringList> InformationArrayValue_list;
    QList<QStringList> AttachInfoValue_list;
    QList<QStringList> RotatedRectValue_list;

    //记录内容对象（主键）
    QList<QStringList> InformationArrayKey_list;
    QList<QStringList> AttachInfoKey_list;
    QList<QStringList> RotatedRectKey_list;

    for(int i = 0 ; i < filePathlistlog.count(); i++){//读取LOG文件
        qDebug()<<filePathlistlog.at(i);
        DisPlayTip(filePathlistlog.at(i));
        QFile file(filePathlistlog.at(i));//打开当前Result.json文件
        if(!file.open(QIODevice::ReadOnly |QIODevice::Text) ){//以读的方式打开文件
            qDebug()<<file.errorString();
            DisPlayTip(file.errorString());
            return false;
        }

        QTextStream in(&file);
        in.setCodec("GBK"); // 设置文件的编码格式为GBK
        QString temp = in.readAll();
        if(!temp.isEmpty()){
             QJsonParseError jsonError;
             QJsonDocument jsonDoc(QJsonDocument::fromJson(temp.toUtf8(), &jsonError));

             if(jsonError.error != QJsonParseError::NoError){
                 qDebug() << "json error!" << jsonError.errorString();
                 DisPlayTip("no json error!");
                 return false;
             }
             QJsonObject jsonObject = jsonDoc.object();
             QStringList keys = jsonObject.keys();
             for(int i = 0; i < keys.size(); i++){
                 qDebug() << "key" << i << " is:" << keys.at(i);
             }

             MainValuelist.clear();
             MainKeylist.clear();
             for(int k = 0; k < Set_MainKeylist.count(); k++ ){
                 if (jsonObject.contains(Set_MainKeylist.at(k))){
                     QJsonValue main_value = jsonObject.take(Set_MainKeylist.at(k));
                     if (main_value.isDouble()){
                       int  nmainvalue = main_value.toVariant().toInt();
                       qDebug()<<Set_MainKeylist.at(k)<< " = "<< nmainvalue;

                       MainValuelist.push_back(QString::number(nmainvalue));
                       MainKeylist.push_back(Set_MainKeylist.at(k));
                     }
                     if (main_value.isBool()) {
                         bool bmainvalue = main_value.toVariant().toBool();
                         qDebug()<<Set_MainKeylist.at(k)<< " = "<< bmainvalue;
                        MainValuelist.push_back(QString::number(bmainvalue));
                        MainKeylist.push_back(Set_MainKeylist.at(k));
                     }
                     if (main_value.isString()){
                       QString  srtmainvalue = main_value.toVariant().toString();
                        qDebug()<<Set_MainKeylist.at(k)<< " = "<< srtmainvalue;
                       MainValuelist.push_back(srtmainvalue);
                       MainKeylist.push_back(Set_MainKeylist.at(k));
                     }
                     if(main_value.isArray()){

                         InformationArrayValue_list.clear();//清理 上一个 文件的数组list
                         InformationArrayKey_list.clear();
                         RotatedRectValue_list.clear();
                         RotatedRectKey_list.clear();
                         AttachInfoValue_list.clear();
                         AttachInfoKey_list.clear();

                         QJsonArray array = main_value.toArray();
                         for(int i=0;i<array.size();i++){

                             QJsonValue defectArray = array.at(i);
                             QJsonObject defect = defectArray.toObject();

                             InformationArrayValuelist.clear();//清理 数组、内容对象 list
                             InformationArrayKeylist.clear();

                             for(int n = 0; n < Set_InformationArrayKeylist.count(); n++){

                                 QJsonValue InformationArrayValue = defect.value(Set_InformationArrayKeylist.at(n));
                                 if(InformationArrayValue.isObject()){

                                     if("AttachInfo" == Set_InformationArrayKeylist.at(n)){
                                         QJsonObject AttachInfoObject = InformationArrayValue.toObject();
                                          QStringList AttachObjectlist = AttachInfoObject.keys();
                                          AttachInfoValuelist.clear();//清理AttachInfo的内容
                                          AttachInfoKeylist.clear();//清理AttachInfo的對象
                                          if(!AttachObjectlist.isEmpty()){
                                              for(int l = 0; l < Set_AttachInfoKeylist.count(); l++){
                                                 AttachInfoValuelist.push_back(AttachInfoObject[Set_AttachInfoKeylist.at(l)].toString());//添加数据
                                                 AttachInfoKeylist.push_back(Set_AttachInfoKeylist.at(l));//添加内容对象
                                                 qDebug()<<Set_AttachInfoKeylist.at(l)<< " = "<<AttachInfoObject[Set_AttachInfoKeylist.at(l)].toString();
                                              }
                                          }
                                          AttachInfoValue_list.push_back(AttachInfoValuelist);
                                          AttachInfoKey_list.push_back(AttachInfoKeylist);
                                     }

                                     if("RotatedRect" == Set_InformationArrayKeylist.at(n)){
                                         QJsonObject RotatedRectObject = InformationArrayValue.toObject();
                                          QStringList RotatedRectObjectlist = RotatedRectObject.keys();
                                          RotatedRectValuelist.clear();//清理 RotatedRect 的内容
                                          RotatedRectKeylist.clear();//清理 RotatedRect 的對象
                                          if(!RotatedRectObjectlist.isEmpty()){
                                              for(int l = 0; l < Set_RotatedRectKeylist.count(); l++){
                                                 RotatedRectValuelist.push_back(QString::number(RotatedRectObject[Set_RotatedRectKeylist.at(l)].toInt()));//添加数据
                                                 RotatedRectKeylist.push_back(Set_RotatedRectKeylist.at(l));//添加内容对象
                                                 qDebug()<<Set_RotatedRectKeylist.at(l)<< " = "<<RotatedRectObject[Set_RotatedRectKeylist.at(l)].toInt();
                                              }
                                          }
                                          RotatedRectValue_list.push_back(RotatedRectValuelist);
                                          RotatedRectKey_list.push_back(RotatedRectKeylist);
                                     }
                                 }

                                 if(InformationArrayValue.isDouble()){//如果是數字
                                    InformationArrayValuelist.push_back(QString::number(defect[Set_InformationArrayKeylist.at(n)].toInt()));//添加数据
                                    InformationArrayKeylist.push_back(Set_InformationArrayKeylist.at(n));//添加内容对象
                                     qDebug()<<Set_InformationArrayKeylist.at(n)<< " = "<<defect[Set_InformationArrayKeylist.at(n)].toInt();
                                 }
                                 if(InformationArrayValue.isString()){//如果是字符
                                    InformationArrayValuelist.push_back(defect[Set_InformationArrayKeylist.at(n)].toString());//添加数据
                                    InformationArrayKeylist.push_back(Set_InformationArrayKeylist.at(n));//添加内容对象
                                    qDebug()<<Set_InformationArrayKeylist.at(n)<< " = "<<defect[Set_InformationArrayKeylist.at(n)].toString();
                                 }
                            }
                            InformationArrayValue_list.push_back(InformationArrayValuelist);
                            InformationArrayKey_list.push_back(InformationArrayKeylist);

                        }
                    }
                 }
             }

             mainValue_list.push_back(MainValuelist);
             mainKey_list.push_back(MainKeylist);

             informationArrayValue_vectorlist.push_back(InformationArrayValue_list);
             AttachInfoValue_vectorlist.push_back(AttachInfoValue_list);
             RotatedRectValue_vectorlist.push_back(RotatedRectValue_list);

             informationArrayKey_vectorlist.push_back(InformationArrayKey_list);
             AttachInfoKey_vectorlist.push_back(AttachInfoKey_list);
             RotatedRectKey_vectorlist.push_back(RotatedRectKey_list);

          }
        file.close();//关闭文件
     }
    return true;
}


void ResultDataAnalyze::ResultDataCompared()
{
    QString tip ;
    QString ReCt ;
    QStringList list1;
    QStringList list2;
    QString filename1;
    QString filename2;
    QString strname1 ;
    QString strname2 ;

    QStringList mainKeylist1   ;
    QStringList mainValuelist1 ;
    QStringList mainValuelist2 ;

    QList<QStringList> inforarrayKey ;
    QList<QStringList> inforarray1   ;
    QList<QStringList> inforarray2   ;
    QStringList  inforarrayKeylist   ;
    QStringList  inforarrayValuelist1;
    QStringList  inforarrayValuelist2;

    QList<QStringList> AttachInfoKey   ;
    QList<QStringList> AttachInfoValue1 ;
    QList<QStringList> AttachInfoValue2  ;
    QStringList AttachInfoKeylist;
    QStringList AttachInfo1;
    QStringList AttachInfo2;

    QList<QStringList> RotatedRectKey   ;
    QList<QStringList> RotatedRectValue1 ;
    QList<QStringList> RotatedRectValue2 ;
    QStringList RotatedRectKeylist;
    QStringList RotatedRect1   ;
    QStringList RotatedRect2  ;

    listcountlist.clear();
    rowcountlist.clear();
    counstrlist.clear();

    try {
        int dircount = 0;
        int nRestul = 0;
        int i = 0;
        int row = 1;
        int defectcount;
        int H = 0;
        for(int C = 0; C < NameCountlist.count(); C++){
            dircount += NameCountlist.at(C).toInt();
            qDebug()<<dircount;
            qDebug()<< i;
            for( int k = i + 1 ; k < dircount ; k++){

            row = row + 1;
            defectcount = 0;
            tip = QString::number(i);
            list1.clear();
            list2.clear();
            filename1 = Namelist.at(i);
            filename2 = Namelist.at(k);
            list1     = filename1.split("_");
            list2     = filename2.split("_");
            filename1 = list1.at(1);
            filename2 = list2.at(1);
            recrdlog->strlog +=Namelist.at(i) + " ** " + Namelist.at(k) + "\n";
            qDebug()<<Namelist.at(i)<< i << k ;
            qDebug()<<Namelist.at(k)<< i << k ;
            setData(0,row, Namelist.at(i) + ";" + Namelist.at(k));
            if(filename1 != filename2)
            {
                DisPlayTip(filename1+" ** "+ tip +" "+ filename2 +" 不匹配 ");
//                return;
            }

             mainKeylist1.clear();
             mainValuelist1.clear();
             mainValuelist2.clear();
             mainKeylist1   = Q_mainKey_list.at(i);
             mainValuelist1 = Q_mainValue_list.at(i);
             mainValuelist2 = Q_mainValue_list.at(k);
            H = 2;
            if(mainValuelist1.count() == mainValuelist2.count()){
                for(int x = 0;  x < mainValuelist1.count(); x++){
                    strname1 = mainValuelist1.at(x);
                    strname2 = mainValuelist2.at(x);
                    if(strname1 != strname2){

                        DisPlayTip( filename1+" ** "+ tip +" "+ mainKeylist1.at(x) +" 不匹配 ");
                        setData(H + x,row, strname1 + ";" + strname2);
                        defectcount++;
                        if("Result" == mainKeylist1.at(x)){
                             nRestul = 1;
                        }
                    }
                }
                if(1 == nRestul)
                    DisPlayTip("Result 数量不一致");
                    setData(1,row, QString::number(defectcount));
                    continue;
            }
            inforarrayKey.clear();
            inforarray1.clear();
            inforarray2.clear();
            inforarrayKey = Q_informationArrayKey_vectorlist.at(i);
            inforarray1   = Q_informationArrayValue_vectorlist.at(i);
            inforarray2   = Q_informationArrayValue_vectorlist.at(k);
            H = 2 + mainKeylist1.count();
            if(inforarray1.count() == inforarray2.count()){

                for(int D = 0 ; D < inforarray1.count(); D++){
                    ReCt = QString::number(D);

                    inforarrayKeylist.clear();
                    inforarrayValuelist1.clear();
                    inforarrayValuelist2.clear();
                    inforarrayKeylist    = inforarrayKey.at(D);
                    inforarrayValuelist1 = inforarray1.at(D);
                    inforarrayValuelist2 = inforarray2.at(D);

                        if(inforarrayValuelist1.count() == inforarrayValuelist2.count()){
                               for(int x = 0; x < inforarrayValuelist1.count(); x++){
                                   strname1 = inforarrayValuelist1.at(x);
                                   strname2 = inforarrayValuelist2.at(x);

                                       if(strname1 != strname2){
                                           DisPlayTip(filename1+" ** "+ tip + " " + QString::number(H + x) +" "+ inforarrayKeylist.at(x) +" 不匹配 ");
                                           setData(H + x,row, ReCt+" -> "+strname1 + ";" + strname2+"\n");
                                           defectcount++;
                                        }
                                }
                          }
                    }
             }
//             else {
//                //報錯缺陷信息數目不一致
//                DisPlayTip("inforarray 差异数量不一致");
//                setData(1,row, QString::number(defectcount));
//                continue;
//             }

            AttachInfoKey.clear();
            AttachInfoValue1.clear();
            AttachInfoValue2.clear();
            AttachInfoKey      = Q_AttachInfoKey_vectorlist.at(i);
            AttachInfoValue1   = Q_AttachInfoValue_vectorlist.at(i);
            AttachInfoValue2   = Q_AttachInfoValue_vectorlist.at(k);
             H = 2 + mainKeylist1.count() + inforarrayKeylist.count();
            if(AttachInfoValue1.count() == AttachInfoValue2.count()){
                for(int D = 0 ; D < AttachInfoValue1.count(); D++){
                    ReCt = QString::number(D);

                    AttachInfoKeylist.clear();
                    AttachInfo1.clear();
                    AttachInfo2.clear();
                    AttachInfoKeylist = AttachInfoKey.at(D);
                    AttachInfo1       = AttachInfoValue1.at(D);
                    AttachInfo2       = AttachInfoValue2.at(D);

                    if(AttachInfo1.count() == AttachInfo2.count()){
                        for(int x = 0; x < AttachInfo1.count(); x++){
                            strname1 = AttachInfo1.at(x);
                            strname2 = AttachInfo2.at(x);

                            if(strname1 != strname2){
                                DisPlayTip(filename1+" ** "+ tip + " " + QString::number(H + x) +" "+ AttachInfoKeylist.at(x) +" 不匹配 ");
                                setData(H + x,row, ReCt+" -> " + strname1 + ";" + strname2+"\n");
                                defectcount++;
                            }
                       }
                  }
               }
            }
//            else {
//                //報錯缺陷信息數目不一致
//                DisPlayTip("AttachInfo 差异数量不一致");
//                setData(1,row, QString::number(defectcount));
//                continue;
//            }

            RotatedRectKey.clear();
            RotatedRectValue1.clear();
            RotatedRectValue2.clear();
            RotatedRectKey      = Q_RotatedRectKey_vectorlist.at(i);
            RotatedRectValue1   = Q_RotatedRectValue_vectorlist.at(i);
            RotatedRectValue2   = Q_RotatedRectValue_vectorlist.at(k);
             H = 2 + mainKeylist1.count() + inforarrayKeylist.count() + AttachInfoKeylist.count();
            if(RotatedRectValue1.count() == RotatedRectValue2.count()){

                   for(int D = 0 ; D < RotatedRectValue1.count(); D++){
                        ReCt = QString::number(D);
                        RotatedRectKeylist.clear();
                        RotatedRect1.clear();
                        RotatedRect2.clear();
                        RotatedRectKeylist = RotatedRectKey.at(D);
                        RotatedRect1       = RotatedRectValue1.at(D);
                        RotatedRect2       = RotatedRectValue2.at(D);
                     if(RotatedRect1.count() == RotatedRect2.count()){
                         for(int x = 0; x < RotatedRect1.count(); x++){

                             strname1 = RotatedRect1.at(x);
                             strname2 = RotatedRect2.at(x);
                             if(strname1 != strname2){
                                 DisPlayTip(filename1+" ** "+ tip + " " + QString::number(H + x) +" "+ RotatedRectKeylist.at(x) +" 不匹配 ");
                                 setData(H + x,row, ReCt+" -> " + strname1 + ";" + strname2+"\n");
                                 defectcount++;
                             }
                         }
                     }
                }
            }
//            else {
//                //報錯缺陷信息數目不一致
//                DisPlayTip("RotatedRect 差异数量不一致");
//                setData(1,row, QString::number(defectcount));
//                continue;
//            }
            setData(1,row, QString::number(defectcount));

          }
        i = dircount ;
      }

    } catch (...){
        DisPlayTip("**对比出现错误**");
        return;
    }

}

void ResultDataAnalyze::setExcelTebel(QList<QStringList> &mainKey_list,QVector<QList<QStringList>> &informationArrayKey_vectorlist,
QVector<QList<QStringList>> &AttachInfoKey_vectorlist, QVector<QList<QStringList>> &RotatedRectKey_vectorlist)
{
    QStringList mainKeylist                = mainKey_list.at(0);
    QList<QStringList> inforarrayKey       = informationArrayKey_vectorlist.at(0);
    QList<QStringList> AttachInfoKey       = AttachInfoKey_vectorlist.at(0);
    QList<QStringList> RotatedRectKey      = RotatedRectKey_vectorlist.at(0);

    ExcelKeylist.clear();//清空
    ExcelKeylist<<"测试文件名"<<"差异数量";
    for(int i = 0 ; i < mainKeylist.count();i++){
        ExcelKeylist.push_back(mainKeylist.at(i));
    }

    for(int i = 0 ; i < inforarrayKey.count();i++){
        QStringList inforarrayKeylist  = inforarrayKey.at(i);
        if(!inforarrayKeylist.isEmpty()){
            for(int k = 0; k < inforarrayKeylist.count();k++){
                ExcelKeylist.push_back(inforarrayKeylist.at(k));
            }
            break;
        }else{
            continue;
        }
    }
    for(int i = 0 ; i < AttachInfoKey.count();i++){
        QStringList AttachInfoKeylist  = AttachInfoKey.at(i);
        if(!AttachInfoKeylist.isEmpty()){
            for(int k = 0; k < AttachInfoKeylist.count();k++){
                ExcelKeylist.push_back(AttachInfoKeylist.at(k));
            }
            break;
        }else{
            continue;
        }
    }
    for(int i = 0 ; i < RotatedRectKey.count();i++){
        QStringList RotatedRectKeylist  = RotatedRectKey.at(i);
        if(!RotatedRectKeylist.isEmpty()){
            for(int k = 0; k < RotatedRectKeylist.count();k++){
                ExcelKeylist.push_back(RotatedRectKeylist.at(k));
            }
            break;
        }else{
            continue;
        }
    }
}

void ResultDataAnalyze::SetFileTabel()
{
    Set_MainKeylist.clear();
    Set_InformationArrayKeylist.clear();
    Set_AttachInfoKeylist.clear();
    Set_RotatedRectKeylist.clear();

    QString  setfilePath = filePath +"/SetFile/";
     qDebug()<<setfilePath;
     Setdir = new QDir(setfilePath);
     QStringList l1=Setdir->entryList(QDir::Files,QDir::Name); //获取当前文件夹的数量
     for (int i = 0;i < l1.size(); i++){
         qDebug()<<l1.at(0);
         SetFileName = l1.at(0);
     }
     qDebug()<<setfilePath+SetFileName;
     //读取json文件的内容
     QFile file(setfilePath+SetFileName);//打开当前Result.json文件
     if(!file.open(QIODevice::ReadOnly |QIODevice::Text) )//以读的方式打开文件
     {
         qDebug()<<file.errorString();
         DisPlayTip(file.errorString());
         return;
     }
     QTextStream in(&file);
     in.setCodec("GBK"); // 设置文件的编码格式为GBK
     QString temp = in.readAll();
     if(!temp.isEmpty())
      {
          QJsonParseError jsonError;
          QJsonDocument jsonDoc(QJsonDocument::fromJson(temp.toUtf8().data(), &jsonError));

          if(jsonError.error != QJsonParseError::NoError)
          {
              qDebug() << "json error!" << jsonError.errorString();
              DisPlayTip("no 1 json error!");
              return;
          }

          QJsonObject jsonObject = jsonDoc.object();
          QStringList keys = jsonObject.keys();
          for(int i = 0; i < keys.size(); i++)
          {
              qDebug() << "key" << i << " is:" << keys.at(i);
          }

          Set_MainKeylist.clear();
          if (jsonObject.contains("main"))
          {
              QJsonValue main_value = jsonObject.take("main");
              if (main_value.isObject()){
                  QJsonObject mainObject = main_value.toObject();
                  QStringList keys = mainObject.keys();
                   qDebug()<<"main=======";
                  for(int i = 0 ; i < keys.size(); i++)
                  {
                      QString keystr = QString::number(i);
                    Set_MainKeylist.push_back(mainObject[keystr].toString()) ;
                    qDebug()<<mainObject[keystr].toString();
                  }
              }
          }
          Set_InformationArrayKeylist.clear();
          if (jsonObject.contains("InformationArray"))
          {
              QJsonValue InformationArray_value = jsonObject.take("InformationArray");
              if (InformationArray_value.isObject()){
                  QJsonObject InformationArrayObject = InformationArray_value.toObject();
                  QStringList keys = InformationArrayObject.keys();
                   qDebug()<<"InformationArray=====";
                  for(int i = 0 ; i < keys.size(); i++)
                  {
                      QString keystr = QString::number(i);
                      Set_InformationArrayKeylist.push_back(InformationArrayObject[keystr].toString()) ;
                    qDebug()<<InformationArrayObject[keystr].toString();
                  }
              }
          }

          Set_AttachInfoKeylist.clear();
          if (jsonObject.contains("AttachInfo"))
          {
              QJsonValue AttachInfo_value = jsonObject.take("AttachInfo");
              if (AttachInfo_value.isObject()){
                  QJsonObject AttachInfoObject = AttachInfo_value.toObject();
                  QStringList keys = AttachInfoObject.keys();
                   qDebug()<<"AttachInfo=====";
                  for(int i = 0 ; i < keys.size(); i++)
                  {
                      QString keystr = QString::number(i);
                      Set_AttachInfoKeylist.push_back(AttachInfoObject[keystr].toString()) ;
                    qDebug()<<AttachInfoObject[keystr].toString();
                  }
              }
          }
          Set_RotatedRectKeylist.clear();
          if (jsonObject.contains("RotatedRect"))
          {
              QJsonValue RotatedRect_value = jsonObject.take("RotatedRect");
              if (RotatedRect_value.isObject()){
                  QJsonObject RotatedRectObject = RotatedRect_value.toObject();
                  QStringList keys = RotatedRectObject.keys();
                   qDebug()<<"RotatedRect=====";
                  for(int i = 0 ; i < keys.size(); i++)
                  {
                      QString keystr = QString::number(i);
                      Set_RotatedRectKeylist.push_back(RotatedRectObject[keystr].toString()) ;
                      qDebug()<<RotatedRectObject[keystr].toString();
                  }
              }
          }

          if(Set_MainKeylist.isEmpty() || Set_InformationArrayKeylist.isEmpty()
             || Set_AttachInfoKeylist.isEmpty() || Set_RotatedRectKeylist.isEmpty())
          {
              DisPlayTip("配置文件内容有缺，请按要求配置好文件!");
          }else {
              DisPlayTip("配置完成!");
          }
     }

     return;
}

void ResultDataAnalyze::on_inforEdit_textChanged()
{
    ui->inforEdit->moveCursor(QTextCursor::End);
}

void ResultDataAnalyze::RecordDataToExcel(QString Path)
{
    try {

    //根据当前时间创建ex表名
    qDebug()<<Path;
    QDir *DataFile = new QDir;
    bool exist = DataFile->exists("ComparedData");
    if(!exist)
    {
        bool isok = DataFile->mkdir("ComparedData"); // 新建文件夹
            if(!isok){
               DisPlayTip("创建文件夹AnalyzeData失败") ;//提示
                return;
             }
    }
    QDateTime current_date_time =QDateTime::currentDateTime();
    QString current_date =current_date_time.toString("yyyyMMddhhmmss");
    QString absFilePath =Path + "/ComparedData/"+current_date+".xlsx";
    absFilePath = absFilePath.replace("/", "\\");
     qDebug()<<absFilePath;
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

      qDebug()<<"设置标题";
     //设置标题
     for(int i = 0; i < cellrowlist.count(); i++){
        QString str = cellrowlist.at(i)+QString::number(1);
        QString Tostr = ExcelKeylist.at(i);
        qDebug()<<str;

       QAxObject * cell = worksheet->querySubObject("Range(QVariant, QVariant)",str);//获取单元格
        cell->dynamicCall("SetValue(const QVariant&)",QVariant(Tostr));//设置单元格的值
        cell->setProperty("HorizontalAlignment", -4108); //左对齐（xlLeft）：-4131  居中（xlCenter）：-4108  右对齐（xlRight）：-415

     }
     DisPlayTip("开始写入");//提示

     for (int i = 0 ; i < listcountlist.count(); i++) {
         listcount = listcountlist.at(i).toInt();
         rowcount  = rowcountlist.at(i).toInt();
         StrData   = counstrlist.at(i);

         counstr = cellrowlist.at(listcount)+QString::number(rowcount);
         Colorcell = worksheet->querySubObject("Range(QVariant, QVariant)",counstr);//获取单元格
         var = Colorcell->dynamicCall("Value");
         StrData = var.toString() + StrData;
         Colorcell->dynamicCall("SetValue(const QVariant&)",QVariant(StrData));//设置单元格的值
         Colorcell->setProperty("HorizontalAlignment", -4108); //左对齐（xlLeft）：-4131  居中（xlCenter）：-4108  右对齐（xlRight）：-4152

     }
     workbook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(absFilePath)); //保存到filepath
     workbook->dynamicCall("Close (Boolean)", false);  //关闭文件
     excel->dynamicCall("Quit(void)");  //退出
     DisPlayTip("输出文件："+absFilePath) ;//提示
     delete excel;
     excel = nullptr;

    } catch (...){
    DisPlayTip("** 写入Excel出现错误 **");
    return;
    }

}

//初始化log
void ResultDataAnalyze::initlog()
{
   if(false == recrdlog->bInitLog){
        recrdlog->init("ResultDataAnalyzeLog");
   }
}

