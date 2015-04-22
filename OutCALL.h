#ifndef OUTCALL_H
#define OUTCALL_H

#include "ContactManager.h"
#include "AsteriskManager.h"

#include <QWidget>
#include <QMap>
#include <QSystemTrayIcon>

class QSystemTrayIcon;
class QMenu;
class DebugInfoDialog;
class QTcpSocket;
class AboutDialog;
class CallHistoryDialog;
class SettingsDialog;
class PlaceCallDialog;

class OutCall : public QWidget
{
public:
    OutCall();
    ~OutCall();

    void show();  

protected:

    void automaticlySignIn();
    void createContextMenu();

protected slots:
    void signInOut();
    void onSettingsDialog();
    void onAboutDialog();
    void onDebugInfo();
    void onActiveCalls();
    void onPlaceCall();
    void onSyncOutlook();
    void onCallHistory();
    void onActivated(QSystemTrayIcon::ActivationReason reason);
    void displayError(QAbstractSocket::SocketError socketError, const QString &msg);
    void onSyncing(bool status);
    void onAddContact();
    void onStateChanged(AsteriskManager::AsteriskState state);

    void onMessageReceived(const QString &message);
    void onCallDeteceted(const QMap<QString, QVariant> &call, AsteriskManager::CallState state);
    void onCallReceived(const QMap<QString, QVariant> &call);

    void onOnlineHelp();
    void onOnlinePdf();
    void close();
    void changeIcon();

private:
    QMenu *m_menu;
    QAction *m_signIn;
    QAction *m_placeCall;

    QSystemTrayIcon *m_systemTryIcon;
    DebugInfoDialog *m_debugInfoDialog;
    SettingsDialog *m_settingsDialog;
    AboutDialog *m_aboutDialog;
    CallHistoryDialog *m_callHistoryDialog;
    PlaceCallDialog *m_placeCallDialog;
    ContactManager m;
    QTimer m_timer;
    bool m_switch;
};

#endif // OUTCALL_H
