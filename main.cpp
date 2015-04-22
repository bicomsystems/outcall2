#include "OutCALL.h"
#include "Global.h"
#include "LocalServer.h"
#include "ContactManager.h"
#include "Notifier.h"

#include "Windows.h"

#include <QApplication>
#include <QMessageBox>
#include <QLocalSocket>
#include <QTranslator>
#include <QDir>


int main(int argc, char *argv[])
{
    QApplication app(argc, argv);
    app.setQuitOnLastWindowClosed(false);
    app.setApplicationName(APP_NAME);
    app.setApplicationVersion("2.0");
    app.setOrganizationName(ORGANIZATION_NAME);

    bool bCallRequest = false;

    g_LanguagesPath         = QApplication::applicationDirPath() + "/lang";
    g_AppSettingsFolderPath = QDir::homePath() + "/OutCALL";
    g_AppDirPath            = QApplication::applicationDirPath();

    if (argc==2 && QString(argv[1]) == "installer")
    {
        // Setup log file paths
        QDir().mkpath(g_AppSettingsFolderPath);
        QFile(g_AppSettingsFolderPath + "/outcall.log").remove();

        if (global::IsOutlookInstalled())
        {
            if (global::EnableOutlookIntegration(true))
                global::log("Outlook plugin installed successfully.", LOG_INFORMATION);
            else
                global::log("Could not install Outlook plugin.", LOG_ERROR);
        }
        else
        {
            global::log("Outlook was not detected.", LOG_INFORMATION);
        }
        return 0;
    }

    if (argc == 2)
    {
        bCallRequest = QString(argv[1]).contains("Dial#####");
    }

    if (bCallRequest)
    {
        QStringList arguments = QString(argv[1]).split("#####");
        QString contactName = arguments[1];
        QString numbers = QString(arguments[2]).replace("outcall://", "");

        contactName.replace("&&&", " ");
        numbers.replace("&&&", " ");

        QLocalSocket s;
        s.connectToServer(LOCAL_SERVER_NAME);

        global::log("MAIN LOCAL", LOG_INFORMATION);

        QString msg = QObject::tr("It appears that %1 is not running.\n" \
                                  "Note for Windows 7/Vista users: please make sure that %2 and Outlook are running under same level of privileges. " \
                                  "Either both elevated (Run as Administrator option) or non-elevated.").arg(APP_NAME).arg(APP_NAME);

        if (!s.waitForConnected(2000))
        {
            MsgBoxInformation(msg);
        }
        else
        {
            if (!s.waitForReadyRead(2000)) // wait for "OK" from the local server
            {
                MsgBoxInformation(QObject::tr("Timeout on local socket. Maybe %1 is not running?").arg(APP_NAME));
            }
            else
            {
                QByteArray socket_data = QString("outlook_call %1 %2\n").arg(QString(contactName.toLatin1().toBase64())).
                                         arg(QString(numbers.toLatin1().toBase64())).toLatin1();
                s.write(socket_data);

                if (!s.waitForBytesWritten(2000))
                    MsgBoxError(QObject::tr("Failed local socket write()."));
            }
        }
        s.disconnectFromServer();
        s.close();

        return 0;
    }

    Notifier notifier;
    QString lang = global::getSettingsValue("language", "general").toString();
    QTranslator translator;
    if (!lang.isEmpty())
    {
        if (translator.load(QString("%1/%2.lang").arg(g_LanguagesPath).arg(lang)))
        {
            qApp->installTranslator(&translator);
        }
        else
        {
            global::setSettingsValue("language", "", "general");
            MsgBoxError(QObject::tr("Failed to load language file."));
        }
    }

    QString username  = global::getSettingsValue("username", "settings").toString();
    QByteArray secret = global::getSettingsValue("password", "settings").toByteArray();
    AsteriskManager manager(username, QString(QByteArray::fromBase64(secret)));

    OutCall outcall;
    outcall.show();

    LocalServer localServer;
    return app.exec();
}
