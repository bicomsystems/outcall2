#include "OutCALL.h"
#include "DebugInfoDialog.h"
#include "SettingsDialog.h"
#include "AboutDialog.h"
#include "CallHistoryDialog.h"
#include "Global.h"
#include "ContactManager.h"
#include "PlaceCallDialog.h"
#include "AsteriskManager.h"
#include "Notifier.h"
#include "PopupWindow.h"

#include <QMenu>
#include <QDebug>
#include <QTcpSocket>
#include <QApplication>
#include <QDesktopServices>
#include <QUrl>
#include <QVariantList>
#include <QVariantMap>
#include <QMessageBox>

OutCall::OutCall() :
    QWidget()
{
    m_systemTryIcon       = new QSystemTrayIcon(this);
    m_menu                = new QMenu(this);
    m_settingsDialog      = new SettingsDialog;
    m_aboutDialog         = new AboutDialog;
    m_debugInfoDialog     = new DebugInfoDialog;
    m_callHistoryDialog   = new CallHistoryDialog;
    m_placeCallDialog     = new PlaceCallDialog;

    connect(m_systemTryIcon,    &QSystemTrayIcon::activated,            this, &OutCall::onActivated);
    connect(g_pContactManager,  &ContactManager::syncing,               this, &OutCall::onSyncing);
    connect(g_pAsteriskManager, &AsteriskManager::messageReceived,      this, &OutCall::onMessageReceived);
    connect(g_pAsteriskManager, &AsteriskManager::callDeteceted,        this, &OutCall::onCallDeteceted);
    connect(g_pAsteriskManager, &AsteriskManager::callReceived,         this, &OutCall::onCallReceived);
    connect(g_pAsteriskManager, &AsteriskManager::error,                this, &OutCall::displayError);
    connect(g_pAsteriskManager, &AsteriskManager::stateChanged,         this, &OutCall::onStateChanged);
    connect(&m_timer,           &QTimer::timeout,                       this, &OutCall::changeIcon);

    global::setSettingsValue("InstallDir", g_AppDirPath.replace("/", "\\"));

    createContextMenu();

    m_systemTryIcon->setContextMenu(m_menu);

    QString path(":/images/disconnected.png");
    m_systemTryIcon->setIcon(QIcon(path));

    automaticlySignIn();
}

OutCall::~OutCall()
{
    delete m_settingsDialog;
    delete m_debugInfoDialog;
    delete m_aboutDialog;
    delete m_callHistoryDialog;
    delete m_placeCallDialog;
}

void OutCall::createContextMenu()
{
    // Exit action
    QAction* exitAction = new QAction(tr("Exit"), m_menu);
    connect(exitAction, &QAction::triggered, this, &OutCall::close);

    // Sign In
    m_signIn  = new QAction(tr("Sign In"), m_menu);
    connect(m_signIn, &QAction::triggered, this, &OutCall::signInOut);

    // Add Contacts
    QAction* addContactAction = new QAction(tr("Add Contact"), m_menu);
    connect(addContactAction, &QAction::triggered, this, &OutCall::onAddContact);

    // SettingsDialog
    QAction* settingsAction = new QAction(tr("Settings"), m_menu);
    connect(settingsAction, &QAction::triggered, this, &OutCall::onSettingsDialog);

    QAction* debugInfoAction = new QAction(tr("Debug info"), m_menu);
    connect(debugInfoAction, &QAction::triggered, this, &OutCall::onDebugInfo);

    // About
    QAction* aboutAction = new QAction(tr("About"), m_menu);
    connect(aboutAction, &QAction::triggered, this, &OutCall::onAboutDialog);

    // Help
    QMenu *helpMenu = new QMenu(tr("Help"), m_menu);
    m_menu->addMenu(helpMenu);

    QAction* onlineHelp = new QAction(tr("Online HTML"), helpMenu);
    connect(onlineHelp, &QAction::triggered, this, &OutCall::onOnlineHelp);

    QAction* onlinePdf = new QAction(tr("Online PDF"), helpMenu);
    connect(onlinePdf, &QAction::triggered, this, &OutCall::onOnlinePdf);

    // Call History
    QAction* callHistoryAction = new QAction(tr("Call History"), m_menu);
    connect(callHistoryAction, &QAction::triggered, this, &OutCall::onCallHistory);

    // Import Contacts
    QMenu *importContactsMenu = new QMenu(tr("Sync Contacts"), m_menu);

    QAction* importOutlookAction = new QAction(tr("From Outlook"), importContactsMenu);
    connect(importOutlookAction, &QAction::triggered, this, &OutCall::onSyncOutlook);

    // Place a Call
    m_placeCall = new QAction(tr("Place a Call"), 0);
    QFont font = m_placeCall->font();
    font.setBold(true);
    m_placeCall->setFont(font);
    connect(m_placeCall, &QAction::triggered, this, &OutCall::onPlaceCall);

    helpMenu->addAction(onlineHelp);
    helpMenu->addAction(onlinePdf);

    // Add actions
    m_menu->addAction(aboutAction);
    m_menu->addSeparator();

    m_menu->addAction(settingsAction);
    m_menu->addAction(debugInfoAction);
    m_menu->addSeparator();

    m_menu->addAction(m_placeCall);
    m_menu->addAction(callHistoryAction);
    m_menu->addMenu(importContactsMenu);
    importContactsMenu->addAction(importOutlookAction);
    m_menu->addSeparator();

    m_menu->addAction(m_signIn);
    m_menu->addAction(addContactAction);
    m_menu->addSeparator();

    m_menu->addAction(exitAction);
}

void OutCall::automaticlySignIn()
{
    bool autoSingIn = global::getSettingsValue("auto_sign_in", "general").toBool();
    if(autoSingIn)
        signInOut();
}

void OutCall::signInOut()
{
    if (m_signIn->text() == "Cancel Sign In")
    {
        g_pAsteriskManager->signOut();
        return;
    }

    if (g_pAsteriskManager->isSignedIn() == false)
    {
        QString server = global::getSettingsValue("servername", "settings").toString();
        QString port = global::getSettingsValue("port", "settings", "5038").toString();
        g_pAsteriskManager->signIn(server, port.toUInt());
    }
    else
    {
        g_pAsteriskManager->signOut();
    }
}

void OutCall::onSyncOutlook()
{
    g_pContactManager->refreshContacts();
}

void OutCall::displayError(QAbstractSocket::SocketError socketError, const QString &msg)
{
    switch (socketError)
    {
    case QAbstractSocket::RemoteHostClosedError:
        MsgBoxInformation(tr("The remote host closed the connection."));
        break;
    case QAbstractSocket::HostNotFoundError:
        MsgBoxInformation(tr("The host was not found. Please check the "
                             "host name and port settings."));
        break;
    case QAbstractSocket::ConnectionRefusedError:
        MsgBoxInformation(tr("The connection was refused by the peer. "
                             "Make sure the server is running, "
                             "and check that the host name and port "
                             "settings are correct."));
        break;
    default:
        MsgBoxInformation(msg);
    }
}

void OutCall::onSyncing(bool status)
{
    if (status)
    {
        foreach(QAction *action, m_menu->actions())
        {
            if (action->text() == tr("Sync Contacts"))
            {
                m_systemTryIcon->setToolTip(tr("Syncing contacts"));
                action->setText(tr("Syncing ..."));
                action->setDisabled(true);
            }
        }
    }
    else
    {
        foreach(QAction *action, m_menu->actions())
        {
            if (action->text() == tr("Syncing ..."))
            {
                if (g_pAsteriskManager->isSignedIn())
                    m_systemTryIcon->setToolTip(tr(APP_NAME) + tr(" - ") + tr("Siged In"));
                else
                    m_systemTryIcon->setToolTip(tr(APP_NAME) + tr(" - ") + tr("Not Siged In"));

                action->setText(tr("Sync Contacts"));
                action->setEnabled(true);
            }
        }
    }
}

void OutCall::onAddContact()
{
    g_pContactManager->addOutlookContact();
}

void OutCall::onMessageReceived(const QString &message)
{
    if (m_debugInfoDialog)
        m_debugInfoDialog->updateDebug(message);
}

void OutCall::onCallDeteceted(const QMap<QString, QVariant> &call, AsteriskManager::CallState state)
{
    m_callHistoryDialog->addCall(call, (CallHistoryDialog::Calls)state);
}

void OutCall::onCallReceived(const QMap<QString, QVariant> &call)
{
    QString from            = call.value("from").toString();
    QString callerName      = call.value("callerIDName").toString();
    bool isMinCallerID      = global::getSettingsValue("min_caller_state", "general").toBool();
    bool contactOnInboud    = global::getSettingsValue("contact_inbound", "outlook").toBool();
    bool contactOnUnknown   = global::getSettingsValue("contact_unknown", "outlook").toBool();
    bool isCallerIDUnknown  = true;
    int callerLength        = from.size();

    QList<Contact*> contactList = g_pContactManager->getContacts();
    QString outlookContactName;
    for(int i = 0; i < contactList.size(); ++i)
    {
        Contact *contact = contactList[i];
        QList<QString> numbers  = contact->numbers.values();
        if (numbers.contains(from))
        {
            outlookContactName = contact->name;
            isCallerIDUnknown = false;
            break;
        }
    }

    if (callerName.isEmpty() || callerName == "<unknown>" || !outlookContactName.isEmpty())
    {
        callerName = outlookContactName;
    }

    if (isMinCallerID)
    {
        int length = global::getSettingsValue("min_caller_id", "general").toInt();
        if (callerLength >= length)
        {
            if (callerName.isEmpty() || callerName == "<unknown>")
            {
                PopupWindow::showCallNotification(QString("(Nr: %2)").arg(from));
            }
            else
            {
                PopupWindow::showCallNotification(QString("%1 (Nr: %2)").arg(callerName).arg(from));
            }
        }
    }
    else
    {
        if (callerName.isEmpty() || callerName == "<unknown>")
        {
            PopupWindow::showCallNotification(QString("(Nr: %2)").arg(from));
        }
        else
        {
            PopupWindow::showCallNotification(QString("%1 (Nr: %2)").arg(callerName).arg(from));
        }
    }

    if (contactOnInboud && !isCallerIDUnknown)
    {
        g_pContactManager->viewOutlookContact(callerName, "");
    }
    else if (isCallerIDUnknown && contactOnUnknown)
    {
        g_pContactManager->addOutlookContact(from, callerName);
    }
}

void OutCall::onStateChanged(AsteriskManager::AsteriskState state)
{
    if (state == AsteriskManager::CONNECTED)
    {
        QString path(":/images/connected.png");
        m_systemTryIcon->setIcon(QIcon(path));

        m_signIn->setText(tr("Sign out"));
        PopupWindow::showInformationMessage(tr(APP_NAME), tr("Signed In"));
        m_systemTryIcon->setToolTip(tr(APP_NAME) + tr(" - ") + tr("Siged In"));
        m_placeCall->setEnabled(true);
        m_timer.stop();
    }
    else if (state == AsteriskManager::CONNECTING)
    {
        m_signIn->setText(tr("Cancel Sign In"));
        m_systemTryIcon->setToolTip(tr(APP_NAME) + tr(" - ") + tr("Signing In"));
        m_placeCall->setEnabled(false);
        m_timer.start(500);
    }
    else if (state == AsteriskManager::DISCONNECTED)
    {
        QString path(":/images/disconnected.png");
        m_systemTryIcon->setIcon(QIcon(path));

        m_signIn->setText(tr("&Sign In"));
        m_systemTryIcon->setToolTip(tr(APP_NAME) + tr(" - ") + tr("Not Siged In"));
        m_placeCall->setEnabled(false);
        m_timer.stop();
    }
    else if (state == AsteriskManager::AUTHENTICATION_FAILED)
    {
        QString path(":/images/started.png");
        m_systemTryIcon->setIcon(QIcon(path));

        PopupWindow::showInformationMessage(tr(APP_NAME), tr("Authentication failed"));
        m_systemTryIcon->setToolTip(tr(APP_NAME) + tr(" - ") + tr("Not configured"));
        m_signIn->setText(tr("&Sign In"));
        m_placeCall->setEnabled(false);
        m_timer.stop();
    }
}

void OutCall::changeIcon()
{
    if (m_switch)
    {
        QString path(":/images/connected.png");
        m_systemTryIcon->setIcon(QIcon(path));
        m_switch = false;
    }
    else
    {
        QString path(":/images/disconnected.png");
        m_systemTryIcon->setIcon(QIcon(path));
        m_switch = true;
    }
}

void OutCall::onOnlineHelp()
{
    QString link = "http://www.bicomsystems.com/docs/outcall/1.0/html/";
    QDesktopServices::openUrl(QUrl(link));
}

void OutCall::onOnlinePdf()
{
    QString link = "http://www.bicomsystems.com/docs/outcall/1.0/pdf/outcall.pdf";
    QDesktopServices::openUrl(QUrl(link));
}

void OutCall::onCallHistory()
{
    m_callHistoryDialog->show();
    m_callHistoryDialog->raise();
}

void OutCall::onSettingsDialog()
{
    SettingsDialog dialog;
    dialog.exec();
}

void OutCall::onAboutDialog()
{
    m_aboutDialog->show();
    m_aboutDialog->raise();
}

void OutCall::onDebugInfo()
{
    m_debugInfoDialog->show();
    m_debugInfoDialog->raise();
}

void OutCall::onPlaceCall()
{
    m_placeCallDialog->show();
    m_placeCallDialog->raise();
}

void OutCall::close()
{
    g_pAsteriskManager->signOut();
    QApplication::quit();
}

void OutCall::onActivated(QSystemTrayIcon::ActivationReason reason)
{
    if (reason == QSystemTrayIcon::Trigger)
    {
        m_debugInfoDialog->activateWindow();
        m_aboutDialog->activateWindow();
        m_settingsDialog->activateWindow();
        m_callHistoryDialog->activateWindow();
        g_pContactManager->activateDialog();
        m_placeCallDialog->activateWindow();
    }
    else if (reason == QSystemTrayIcon::DoubleClick)
    {
        m_placeCallDialog->show();
        m_placeCallDialog->activateWindow();
    }
}

void OutCall::show()
{
    m_systemTryIcon->show();
}
