#ifndef LOCALSERVER_H
#define LOCALSERVER_H

#include <QLocalServer>

#define LOCAL_SERVER_NAME	"outcall_ipc_server"

class QMessageBox;

class LocalServer : public QLocalServer
{
    Q_OBJECT
public:
    explicit LocalServer(QObject *parent = 0);
    ~LocalServer();

signals:

private slots:
    void newSocketConnection();
    void readyRead();

private:
    QMessageBox *m_box;
};

#endif // LOCALSERVER_H
