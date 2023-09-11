
#ifndef QDEBUGOUTPUTREDIRECTOR_H
#define QDEBUGOUTPUTREDIRECTOR_H

#include <QObject>
#include <QPlainTextEdit>
#include <QDateTime>

class QDebugOutputRedirector: public QObject
{
    Q_OBJECT
public:
    explicit QDebugOutputRedirector(QPlainTextEdit *textEdit, QObject *parent = nullptr);

private:
    static void customMessageHandler(QtMsgType type, const QMessageLogContext &context, const QString &msg);
    static QPlainTextEdit* m_textEdit; // 静态成员变量声明
};

#endif // QDEBUGOUTPUTREDIRECTOR_H



