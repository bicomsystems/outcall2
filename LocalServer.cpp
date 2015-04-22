#include "LocalServer.h"
#include "Global.h"
#include "ContactManager.h"
#include "OutCALL.h"
#include "AsteriskManager.h"

#include <QTcpSocket>
#include <QLocalSocket>
#include <QStringList>
#include <QFile>
#include <QMessageBox>

LocalServer::LocalServer(QObject *parent) :
    QLocalServer(parent)
{
    removeServer(LOCAL_SERVER_NAME);

    if (!listen(LOCAL_SERVER_NAME))
        global::log(QString("Failed listen() on local server. %1").arg(errorString()), LOG_ERROR);

    connect(this, SIGNAL(newConnection()), this, SLOT(newSocketConnection()));

    m_box = new QMessageBox;
    m_box->setWindowIcon(QIcon(":images/outcall-logo.png"));
}

LocalServer::~LocalServer()
{
    close();
    removeServer(LOCAL_SERVER_NAME);

    delete m_box;
}

void LocalServer::newSocketConnection()
{
    QLocalSocket *s = nextPendingConnection();

    if (!s)
        return;

    s->setProperty("incoming_data", "");
    s->write("OK");

    connect(s, SIGNAL(disconnected()), s, SLOT(deleteLater()));
    connect(s, SIGNAL(readyRead()), this, SLOT(readyRead()));
}

void LocalServer::readyRead()
{
    QLocalSocket *s = (QLocalSocket*)sender();
    QString data = s->property("incoming_data").toString() + s->readAll();

    global::log("LOCAL SERVER", LOG_INFORMATION);

    if (data.endsWith("\n"))
    {
        s->disconnectFromServer();
        QStringList cmd = data.split(" ");
        if (cmd.count()==3)
        {
             if (cmd[0]=="outlook_call")
             {
                QString contactName = QString(QByteArray::fromBase64(cmd[1].toLatin1()));
                QStringList numbers = QString(QByteArray::fromBase64(cmd[2].toLatin1())).split(" | ");
                numbers.removeAll("");

                if (!g_pAsteriskManager->isSignedIn())
                {
                    m_box->setWindowTitle(tr(APP_NAME));
                    m_box->setText(tr("%1 is not logged in.\nPlease log in and try calling the contact again.").arg(APP_NAME));
                    m_box->setIcon(QMessageBox::Warning);
                    m_box->activateWindow();
                    m_box->show();
                    return;
                }
                QHash<QString, QString> numbers_hash;
                QStringList info;

                for (int i = 0; i < numbers.count(); ++i)
                {
                    info = numbers[i].split(":");

                    if (info.count() != 2)
                        continue;

                    numbers_hash.insert(info[0], info[1]);
                }
                g_pContactManager->openContactDialog(contactName, numbers_hash);
            }
        }
    }
    else
    {
        s->setProperty("incoming_data", data);
    }
}
