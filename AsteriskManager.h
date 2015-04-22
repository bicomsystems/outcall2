#ifndef ASTERISKMANAGER_H
#define ASTERISKMANAGER_H

#include <QMap>
#include <QTcpSocket>
#include <QObject>
#include <QTimer>

class AsteriskManager : public QObject
{
    Q_OBJECT
public:
    explicit AsteriskManager(const QString username, const QString secret, QObject *parent = 0);

    ~AsteriskManager();

    enum CallState
    {
        MISSED = 0,
        RECIEVED = 1,
        PLACED = 2
    };

    enum AsteriskState
    {
        CONNECTED,
        CONNECTING,
        DISCONNECTED,
        AUTHENTICATION_FAILED,
        ERROR_ON_CONNECTING
    };

    enum AsteriskVersion
    {
        VERSION_11,
        VERSION_13
    };

    struct Call
    {
        QString chExten;
        QString chType;
        QString exten;
        QString state;
        QString callerName;
        QString destUniqueid;
    };

    void originateCall(QString from, QString to,  QString protocol,
                        QString callerID, QString Context = "default");

    bool isSignedIn() const;
    void signIn(const QString &serverName, const quint16 &port);
    void signOut();
    void reconnect();
    void setAutoSignIn(bool ok);

    void setState(AsteriskState state);

protected:
    void getEventValues(QString eventData, QMap<QString, QString> &map);

    void parseEvent(const QString &event);

    void asterisk_11_eventHandler(const QString &eventData);

    void formatNumber(QString &number);

    void setAsteriskVersion(const QString &msg);

protected slots:
    void onError(QAbstractSocket::SocketError socketError);
    void read();
    void login();
    void onSettingsChange();

signals:
    void messageReceived(const QString &message);
    void authenticationState(bool state);
    void callDeteceted(const QMap<QString, QVariant> &call, CallState state);
    void callReceived(const QMap<QString, QVariant>&);
    void error(QAbstractSocket::SocketError socketError, const QString &msg);
    void stateChanged(AsteriskState state);

private:
    QTcpSocket *m_tcpSocket;
    QString m_eventData;
    bool m_isSignedIn;
    bool m_autoConnectingOnError;
    QMap<QString, int> m_dialedNum;

    QString m_username;
    QString m_secret;
    QString m_server;
    quint16 m_port;

    QTimer m_timer;
    AsteriskState m_currentState;
    bool m_autoSignIn;
    AsteriskVersion m_currentVersion;

    QMap<QString, Call*> m_calls;
};

extern AsteriskManager *g_pAsteriskManager;

#endif // ASTERISKMANAGER_H
