#include "AsteriskManager.h"
#include "Notifier.h"
#include "Global.h"

#include <QDebug>
#include <QTcpSocket>
#include <QTime>
#include <QVariantList>
#include <QVariantMap>
#include <QTimer>

AsteriskManager *g_pAsteriskManager = nullptr;

AsteriskManager::AsteriskManager(const QString username, const QString secret, QObject *parent)
    : QObject(parent),
      m_isSignedIn(false),
      m_autoConnectingOnError(false),
      m_username(username),
      m_secret(secret)
{
    g_pAsteriskManager = this;

    m_tcpSocket = new QTcpSocket(this);
    void (QAbstractSocket:: *sig)(QAbstractSocket::SocketError) = &QAbstractSocket::error;

    connect(m_tcpSocket, &QIODevice::readyRead,       this, &AsteriskManager::read);
    connect(m_tcpSocket, &QAbstractSocket::connected, this, &AsteriskManager::login);
    connect(m_tcpSocket, sig,                         this, &AsteriskManager::onError);
    connect(&m_timer,    &QTimer::timeout,            this, &AsteriskManager::reconnect);

    connect(g_Notifier,  &Notifier::settingsChanged,  this, &AsteriskManager::onSettingsChange);
    m_currentState = DISCONNECTED;
}

AsteriskManager::~AsteriskManager()
{
}

void AsteriskManager::signIn(const QString &serverName, const quint16 &port)
{
    if (!m_isSignedIn && (m_currentState == DISCONNECTED || m_currentState == ERROR_ON_CONNECTING))
    {
        m_server = serverName;
        m_port   = port;
        setState(CONNECTING);
        m_tcpSocket->connectToHost(serverName, port);
    }
}

void AsteriskManager::signOut()
{
    m_tcpSocket->abort();
    m_isSignedIn = false;
    m_autoConnectingOnError = false;
    m_timer.stop();
    setState(DISCONNECTED);
}

void AsteriskManager::onError(QAbstractSocket::SocketError socketError)
{
    if (m_currentState == CONNECTING && !m_autoConnectingOnError)
    {
        QString msg = m_tcpSocket->errorString();
        emit error(socketError, msg);
        setState(DISCONNECTED);
    }
    else if (m_currentState == CONNECTED)
    {
        setState(DISCONNECTED);
        if (!m_timer.isActive() && m_autoSignIn)
        {
            m_timer.start(30000); // 30 sec to reconnect
        }
    }
    setState(ERROR_ON_CONNECTING);
}

void AsteriskManager::reconnect()
{
    if (!m_autoSignIn)
    {
        m_timer.stop();
        m_autoConnectingOnError = false;
        return;
    }

    m_autoConnectingOnError = true;
    signIn(m_server, m_port);
}

void AsteriskManager::setAutoSignIn(bool ok)
{
    m_autoSignIn = ok;
}

void AsteriskManager::setState(AsteriskState state)
{
    if (state == m_currentState)
        return;

    if (state == CONNECTED)
    {
        m_currentState = CONNECTED;
        emit stateChanged(CONNECTED);
    }
    else if (state == CONNECTING)
    {
        m_currentState = CONNECTING;
        emit stateChanged(CONNECTING);
    }
    else if (state == DISCONNECTED)
    {
        m_currentState = DISCONNECTED;
        m_isSignedIn = false;
        emit stateChanged(DISCONNECTED);
    }
    else if (state == AUTHENTICATION_FAILED)
    {
        m_currentState = DISCONNECTED;
        m_isSignedIn = false;
        emit stateChanged(AUTHENTICATION_FAILED);
    }
    else if (state == ERROR_ON_CONNECTING)
    {
        m_currentState = ERROR_ON_CONNECTING;
    }
}

void AsteriskManager::setAsteriskVersion(const QString &msg)
{
    int index = msg.indexOf("/") + 1;
    QString ami = msg.mid(index);

    if (ami.at(0) == '1')
    {
        m_currentVersion = VERSION_11;
    }
    else if (ami.at(0) == '2')
    {
        m_currentVersion = VERSION_13;
    }
}

void AsteriskManager::read()
{
    while(m_tcpSocket->canReadLine())
    {
        QString message = QString::fromUtf8(m_tcpSocket->readLine());

        if (m_isSignedIn == false)
        {
            QString msg = message.trimmed();

            if (msg.contains("Asterisk Call Manager"))
                setAsteriskVersion(msg);

            if (msg == "Message: Authentication accepted")
            {
                m_isSignedIn = true;
                m_autoConnectingOnError = false;
                m_timer.stop();
                setState(CONNECTED);
            }
            else if(msg == "Message: Authentication failed")
            {
                m_timer.stop();
                m_tcpSocket->abort();
                m_isSignedIn = false;
                m_autoConnectingOnError = false;
                setState(AUTHENTICATION_FAILED);
            }
        }
        else
        {
            if (message != "\r\n")
            {
                m_eventData.append(message);
            }
            else if (!m_eventData.isEmpty())
            {
                if (m_currentVersion == VERSION_13)
                {
                    parseEvent(m_eventData);
                }
                else if (m_currentVersion == VERSION_11)
                {
                    asterisk_11_eventHandler(m_eventData);
                }
                m_eventData.clear();
            }
        }
        emit messageReceived(message.trimmed());
    }
}

void AsteriskManager::parseEvent(const QString &eventData)
{
    if (eventData.contains("Event: Newchannel"))
    {
        QMap<QString, QString> eventValues;
        getEventValues(eventData, eventValues);
        QString uniqueid     = eventValues.value("Uniqueid");
        QString callerIDName = eventValues.value("CallerIDName");
        QString ch           = eventValues.value("Channel");

        QRegExp reg("([^/]*)(/)(\\d+)");
        reg.indexIn(ch);
        QString chType = reg.cap(1);
        QString chExten = reg.cap(3);

        if (!global::containsSettingsKey(chExten, "extensions"))
            return;

        QString type = global::getSettingsValue(chExten, "extensions").toString();
        if (type == chType)
        {
            Call *call       = new Call;
            call->callerName = callerIDName;
            call->chType     = chType;
            call->chExten    = chExten;

            m_calls.insert(uniqueid, call);
        }
    }
    else if (eventData.contains("Event: Newexten"))
    {
        QMap<QString, QString> eventValues;
        getEventValues(eventData, eventValues);
        QString exten       = eventValues.value("Exten");
        QString uniqueid    = eventValues.value("Uniqueid");
        QString state       = eventValues.value("ChannelStateDesc");
        QString appData     = eventValues.value("AppData");

        if (exten == "s" || exten == "h" || !m_calls.contains(uniqueid) || exten.isEmpty())
            return;

        if (appData == "(Outgoing Line)")
            return;

        Call *call = m_calls.value(uniqueid);
        call->state = state;

        if (call->exten.isEmpty())
        {
            QString dateTime = QDateTime::currentDateTime().toString();

            call->exten = exten;
            m_calls.insert(uniqueid, call);

            QMap<QString, QVariant> placed;
            placed.insert("from",       call->chExten);
            placed.insert("to",         call->exten);
            placed.insert("protocol",   call->chType);
            placed.insert("date_time",  dateTime);

            QList<QVariant> list = global::getSettingsValue("placed", "calls", QVariantList()).toList();

            if (list.size() >= 50)
                list.removeFirst();

            list.append(QVariant::fromValue(placed));
            global::setSettingsValue("placed", list, "calls");

            emit callDeteceted(placed, PLACED);
        }
    }
    else if(eventData.contains("Event: DialBegin"))
    {
        QMap<QString, QString> eventValues;
        getEventValues(eventData, eventValues);

        const QString channelStateDesc  = eventValues.value("ChannelStateDesc");
        const QString callerIDNum       = eventValues.value("CallerIDNum");
        QString destExten               = eventValues.value("DestExten");
        const QString destChannel       = eventValues.value("DestChannel");
        const QString callerIDName      = eventValues.value("CallerIDName");
        const QString uniqueid          = eventValues.value("Uniqueid");
        QString destProtocol;

        QRegExp reg("([^/]*)(/)(\\d+)");
        reg.indexIn(destChannel);
        destProtocol = reg.cap(1);
        destExten = reg.cap(3);

        if (channelStateDesc == "Ring" || channelStateDesc == "Up")
        {
            if (global::containsSettingsKey(destExten, "extensions"))
            {
                QString protocol        = global::getSettingsValue(destExten, "extensions").toString();
                bool recievedProtocol   = false;

                if (protocol == destProtocol)
                    recievedProtocol = true;

                if (protocol == "SIP" && destProtocol  == "PJSIP")
                    recievedProtocol = true;

                if (recievedProtocol)
                {
                    QMap<QString, QVariant> received;
                    received.insert("from", callerIDNum);
                    received.insert("to", destExten);
                    received.insert("protocol", destProtocol);
                    received.insert("callerIDName", callerIDName);

                    int counter = m_dialedNum.value(uniqueid, 0);
                    counter++;
                    m_dialedNum.insert(uniqueid, counter);

                    if (counter == 1)
                        emit callReceived(received);
                }
            }
        }
    }
    else if(eventData.contains("Event: DialEnd"))
    {
        QMap<QString, QString> eventValues;
        getEventValues(eventData, eventValues);

        QString channelStateDesc = eventValues.value("ChannelStateDesc");
        QString callerIDNum      = eventValues.value("CallerIDNum");
        QString exten            = eventValues.value("Exten");
        QString destChannel      = eventValues.value("DestChannel");
        QString dialStatus       = eventValues.value("DialStatus");
        QString callerIDName     = eventValues.value("CallerIDName");
        QString uniqueid         = eventValues.value("Uniqueid");

        int index                = destChannel.indexOf("/");
        QString destProtocol     = destChannel.mid(0, index);

        QRegExp reg("([^/]*)(/)(\\d+)");
        reg.indexIn(destChannel);
        destProtocol = reg.cap(1);
        exten = reg.cap(3);

        if (channelStateDesc == "Ring" || channelStateDesc == "Up")
        {
            QString protocol = global::getSettingsValue(exten, "extensions").toString();

            bool isProtocolOk = false;
            if (protocol == destProtocol)
            {
                isProtocolOk = true;
            }
            else if (destProtocol == "PJSIP" && protocol == "SIP")
            {
                isProtocolOk = true;
            }

            if (global::containsSettingsKey(exten, "extensions") && isProtocolOk)
            {
                if (dialStatus == "ANSWER")
                {
                    QString dateTime = QDateTime::currentDateTime().toString();

                    QMap<QString, QVariant> received;
                    received.insert("from",         callerIDNum);
                    received.insert("to",           exten);
                    received.insert("protocol",     destProtocol);
                    received.insert("date_time",    dateTime);
                    received.insert("callerIDName", callerIDName);

                    QList<QVariant> list = global::getSettingsValue("received", "calls").toList();

                    if (list.size() >= 50)
                        list.removeFirst();

                    list.append(QVariant::fromValue(received));
                    global::setSettingsValue("received", list, "calls");

                    m_dialedNum.remove(uniqueid);

                    emit callDeteceted(received, RECIEVED);
                }
                else if (dialStatus == "CANCEL" || dialStatus == "BUSY" || dialStatus == "NOANSWER")
                {
                    int counter = 0;
                    if (m_dialedNum.contains(uniqueid))
                    {
                        counter = m_dialedNum.value(uniqueid, 0);
                    }
                    else
                    {
                        return;
                    }

                    if (counter > 1)
                    {
                        counter--;
                        m_dialedNum.insert(uniqueid, counter);
                        return;
                    }
                    m_dialedNum.remove(uniqueid);

                    QString date = QDateTime::currentDateTime().toString();

                    QMap<QString, QVariant> missed;
                    missed.insert("from",           callerIDNum);
                    missed.insert("to",             exten);
                    missed.insert("protocol",       destProtocol);
                    missed.insert("date_time",      date);
                    missed.insert("callerIDName",   callerIDName);

                    QList<QVariant> list = global::getSettingsValue("missed", "calls", QVariantList()).toList();

                    if (list.size() >= 50)
                        list.removeFirst();

                    list.append(QVariant::fromValue(missed));
                    global::setSettingsValue("missed", list, "calls");

                    emit callDeteceted(missed, MISSED);
                }
            }
        }
    }
    else if(eventData.contains("Event: Hangup"))
    {
        QMap<QString, QString> eventValues;
        getEventValues(eventData, eventValues);

        const QString uniqueid = eventValues.value("Uniqueid");
        if (m_calls.contains(uniqueid))
        {
            Call *call = m_calls.value(uniqueid);
            delete call;
            m_calls.remove(uniqueid);
        }
    }
}

void AsteriskManager::asterisk_11_eventHandler(const QString &eventData)
{
    if (eventData.contains("Event: Newchannel"))
    {
        QMap<QString, QString> eventValues;
        getEventValues(eventData, eventValues);

        QString uniqueid         = eventValues.value("Uniqueid");
        QString channel          = eventValues.value("Channel");
        QString calledNum        = eventValues.value("Exten");
        QString state            = eventValues.value("ChannelStateDesc");
        QString callerIDName     = eventValues.value("CallerIDName");

        QRegExp reg("([^/]*)(/)(\\d+)");
        reg.indexIn(channel);

        QString protocol    = reg.cap(1);
        QString exten       = reg.cap(3);

        if (global::containsSettingsKey(exten, "extensions"))
        {
            Call *call          = new Call;
            call->chExten       = exten;
            call->exten         = calledNum;
            call->state         = state;
            call->chType        = protocol;
            call->callerName    = callerIDName;

            m_calls.insert(uniqueid, call);
        }
    }
    else if(eventData.contains("Event: Newstate"))
    {
        QMap<QString, QString> eventValues;
        getEventValues(eventData, eventValues);

        QString channelStateDesc = eventValues.value("ChannelStateDesc");
        QString uniqueid         = eventValues.value("Uniqueid");
        QString channel          = eventValues.value("Channel");
        QString callerIDName     = eventValues.value("ConnectedLineName");
        QString connectedLineNum = eventValues.value("ConnectedLineNum");

        if (!m_calls.contains(uniqueid))
            return;

        QRegExp reg("([^/]*)(/)(\\d+)");
        reg.indexIn(channel);

        QString protocol    = reg.cap(1);
        QString exten       = reg.cap(3);

        bool processEvent = false;
        if (global::containsSettingsKey(exten, "extensions"))
        {
            QString userProtocol = global::getSettingsValue(exten, "extensions").toString();
            processEvent = userProtocol == protocol ? true : false;
        }

        if (exten == connectedLineNum)
            return;

        if (processEvent)
        {
            QString dateTime = QDateTime::currentDateTime().toString();

            if (channelStateDesc == "Ringing")
            {
                QMap<QString, QVariant> received;
                received.insert("from", connectedLineNum);
                received.insert("callerIDName", callerIDName);

                Call *call = m_calls.value(uniqueid);
                call->state      = channelStateDesc;
                call->exten         = connectedLineNum;

                m_calls.insert(uniqueid, call);

                emit callReceived(received);
            }
            if (channelStateDesc == "Up")
            {
                QMap<QString, QVariant> received;
                received.insert("from",         connectedLineNum);
                received.insert("to",           exten);
                received.insert("protocol",     protocol);
                received.insert("date_time",    dateTime);
                received.insert("callerIDName", callerIDName);

                QList<QVariant> list = global::getSettingsValue("received", "calls").toList();

                if (list.size() >= 50)
                    list.removeFirst();

                list.append(QVariant::fromValue(received));
                global::setSettingsValue("received", list, "calls");

                Call *call = m_calls.value(uniqueid);
                call->state      = channelStateDesc;

                m_calls.insert(uniqueid, call);

                emit callDeteceted(received, RECIEVED);
            }

            if (channelStateDesc == "Ring")
            {
                Call *call = m_calls.value(uniqueid);
                call->state      = channelStateDesc;
                QString to       = call->exten;

                m_calls.insert(uniqueid, call);

                QMap<QString, QVariant> placed;
                placed.insert("from", exten);
                placed.insert("to", to);
                placed.insert("protocol", protocol);
                placed.insert("date_time", dateTime);

                QList<QVariant> list = global::getSettingsValue("placed", "calls", QVariantList()).toList();

                if (list.size() >= 50)
                    list.removeFirst();

                list.append(QVariant::fromValue(placed));
                global::setSettingsValue("placed", list, "calls");

                emit callDeteceted(placed, PLACED);
            }
        }
    }
    else if (eventData.contains("Event: Newexten"))
    {
        QMap<QString, QString> eventValues;
        getEventValues(eventData, eventValues);

        QString uniqueid    = eventValues.value("Uniqueid");
        QString extension   = eventValues.value("Extension");

        if (!m_calls.contains(uniqueid))
            return;

        Call *call = m_calls.value(uniqueid);
        QString to = call->exten;
        if (to.isEmpty())
        {
            call->exten = extension;
            QString dateTime = QDateTime::currentDateTime().toString();

            QMap<QString, QVariant> placed;
            placed.insert("from", call->chExten);
            placed.insert("to", call->exten);
            placed.insert("protocol", call->chType);
            placed.insert("date_time", dateTime);

            QList<QVariant> list = global::getSettingsValue("placed", "calls", QVariantList()).toList();

            if (list.size() >= 50)
                list.removeFirst();

            list.append(QVariant::fromValue(placed));
            global::setSettingsValue("placed", list, "calls");

            emit callDeteceted(placed, PLACED);
        }
    }
    else if(eventData.contains("Event: Hangup"))
    {
        QMap<QString, QString> eventValues;
        getEventValues(eventData, eventValues);

        QString uniqueid    = eventValues.value("Uniqueid");
        QString connectedLineName = eventValues.value("ConnectedLineName");

        if (!m_calls.contains(uniqueid))
            return;

        Call *call = m_calls.value(uniqueid);
        QString state = call->state;

        if (state == "Ringing")
        {
            QString dateTime = QDateTime::currentDateTime().toString();

            QMap<QString, QVariant> missed;
            missed.insert("from",           call->exten);
            missed.insert("to",             call->chExten);
            missed.insert("protocol",       call->chType);
            missed.insert("date_time",      dateTime);
            missed.insert("callerIDName",   connectedLineName);

            QList<QVariant> list = global::getSettingsValue("missed", "calls", QVariantList()).toList();

            if (list.size() >= 50)
                list.removeFirst();

            list.append(QVariant::fromValue(missed));
            global::setSettingsValue("missed", list, "calls");

            emit callDeteceted(missed, MISSED);
        }
        delete call;
        m_calls.remove(uniqueid);
    }
}

void AsteriskManager::getEventValues(QString eventData, QMap<QString, QString> &map)
{
    QStringList list = eventData.split("\r\n");
    list.removeLast();
    foreach (QString values, list)
    {
        QStringList c = values.split(": ");
        map.insert(c.at(0), c.at(1));
    }
}

void AsteriskManager::login()
{
    QString login;
    login =  "Action: Login\r\n";
    login += "Username: "        + m_username      + "\r\n";
    login += "Secret: "          + m_secret        + "\r\n";

    m_tcpSocket->write(login.toLatin1().data());
    m_tcpSocket->write("\r\n");
    m_tcpSocket->flush();
}

void AsteriskManager::originateCall(QString from, QString exten, QString protocol,
                            QString callerID, QString context)
{
    QStringList speedDialList = global::getSettingKeys("speed_dial");
    if(speedDialList.contains(exten))
    {
        exten = global::getSettingsValue(exten, "speed_dial").toString();
    }

    formatNumber(exten);

    context = global::getSettingsValue("context", "dial_rules").toString();
    if (context.isEmpty())
        context = "default";

    QString prefix = global::getSettingsValue("prefix", "dial_rules").toString();

    exten = prefix + exten;

    const QString channel = protocol + "/" + from;

    QString result;
    result =  "Action: Originate\r\n";
    result += "Channel: "       + channel       + "\r\n";
    result += "Exten: "         + exten         + "\r\n";
    result += "Context: "       + context       + "\r\n";
    result += "Priority: 1\r\n";
    result += "CallerID: "      + callerID      + "\r\n";

    m_tcpSocket->write(result.toLatin1().data());
    m_tcpSocket->write("\r\n");
    m_tcpSocket->flush();
}

void AsteriskManager::formatNumber(QString &number)
{
    QString formatted_number = number;
    QString international = global::getSettingsValue("international", "dial_rules").toString();

    QVariantList list = global::getSettingsValue("replacement_rules", "dial_rules").toList();

    for (int i = 0; i < list.size(); ++i)
    {
        QMap<QString, QVariant> rules = list[i].toMap();

        int isRegEx         = rules.value("regex").toInt();
        QString replacement = rules.value("replacement").toString();
        QString text        = rules.value("text").toString();

        if (isRegEx)
        {
            formatted_number.replace(QRegExp(text), replacement);
        }
        else
        {
            formatted_number.replace(text, replacement);
        }
    }

    number.clear();

    char ch;
    for (int i = 0; i < formatted_number.length(); ++i)
    {
        ch = formatted_number.at(i).toLatin1();
        if (ch=='+')
            number.append(international);
        else if ((ch>='0' && ch<='9') || ch=='*' || ch=='#')
            number.append(QChar(ch));
        else if (ch=='a' || ch=='b' || ch=='c' || ch=='A' || ch=='B' || ch=='C')
            number.append(QChar('2'));
        else if (ch=='d' || ch=='e' || ch=='f' || ch=='D' || ch=='E' || ch=='F')
            number.append(QChar('3'));
        else if (ch=='g' || ch=='h' || ch=='i' || ch=='G' || ch=='H' || ch=='I')
            number.append(QChar('4'));
        else if (ch=='j' || ch=='k' || ch=='l' || ch=='J' || ch=='K' || ch=='L')
            number.append(QChar('5'));
        else if (ch=='m' || ch=='n' || ch=='o' || ch=='M' || ch=='N' || ch=='O')
            number.append(QChar('6'));
        else if (ch=='p' || ch=='q' || ch=='r' || ch=='s' || ch=='P' || ch=='Q' || ch=='R' || ch=='S')
            number.append(QChar('7'));
        else if (ch=='v' || ch=='u' || ch=='t' || ch=='V' || ch=='U' || ch=='T')
            number.append(QChar('8'));
        else if (ch=='w' || ch=='x' || ch=='y' || ch=='z' || ch=='W' || ch=='X' || ch=='Y' || ch=='Z')
            number.append(QChar('9'));
    }
}

void AsteriskManager::onSettingsChange()
{
    QString server            = global::getSettingsValue("servername", "settings").toString();
    quint16 port              = global::getSettingsValue("port",       "settings").toUInt();
    QString username          = global::getSettingsValue("username",   "settings").toString();
    QByteArray secretInByte   = global::getSettingsValue("password",   "settings").toByteArray();
    QString secret            = QString(QByteArray::fromBase64(secretInByte));

    if(server != m_server || port != m_port ||
            username != m_username || secret != m_secret)
    {
        m_server    = server;
        m_port      = port;
        m_username  = username;
        m_secret    = secret;

        signOut();
        signIn(m_server, m_port);
    }
}

bool AsteriskManager::isSignedIn() const
{
    return m_isSignedIn;
}
