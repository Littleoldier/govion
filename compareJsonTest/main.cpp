#include "mainwindow.h"
#include "QDebugOutputRedirector.h"
#include <QApplication>

//QPlainTextEdit* QDebugOutputRedirector::m_textEdit = nullptr;

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
//    qInstallMessageHandler(myMessageHandler); // 安装自定义消息处理函数

    MainWindow w;
    w.show();
    return a.exec();
}
