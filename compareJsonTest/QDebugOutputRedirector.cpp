#include "QDebugOutputRedirector.h"
#include <QDebug>
#include <iostream>

QPlainTextEdit* QDebugOutputRedirector::m_textEdit = nullptr; // 静态成员变量定义

QDebugOutputRedirector::QDebugOutputRedirector(QPlainTextEdit *textEdit, QObject *parent)
    : QObject(parent)
{
    m_textEdit = textEdit;
    qInstallMessageHandler(customMessageHandler);
}

void QDebugOutputRedirector::customMessageHandler(QtMsgType type, const QMessageLogContext &context, const QString &msg)
{
    // 构造要显示的消息内容
    QByteArray localMsg = msg.toLocal8Bit();
    QString log;
    QDateTime da_time;
    QString time_str = da_time.currentDateTime().toString("yyyy-MM-dd HH:mm:ss");

    switch (type) {
        case QtDebugMsg:
            log = QString("Debug[%1]: %2").arg(time_str).arg(localMsg.constData());
            break;
        case QtWarningMsg:
            log = QString("Warning[%1]: %2").arg(time_str).arg(localMsg.constData());
            break;
        case QtCriticalMsg:
            log = QString("Critical[%1]: %2").arg(time_str).arg(localMsg.constData());
            break;
        case QtFatalMsg:
            log = QString("Fatal[%1]: %2").arg(time_str).arg(localMsg.constData());
            break;
    }

    if (m_textEdit) {
        // 将消息追加到QPlainTextEdit
        m_textEdit->appendPlainText(log);

        // 在控制台打印该消息
        std::cout << log.toStdString() << std::endl;
    }

//        // 写入日志文件
//        QFile file("log.txt");
//        if (file.open(QIODevice::WriteOnly | QIODevice::Append | QIODevice::Text)) {
//            QTextStream stream(&file);
//            stream << log << endl;
//            file.close();
//        }
}
