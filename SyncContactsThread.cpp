#include "SyncContactsThread.h"
#include "ContactManager.h"
#include "Global.h"

#include <QSqlDatabase>
#include <QMutex>
#include <QApplication>
#include <QProcess>
#include <QDir>

#define SQLITE_DB_CONNECTION_NAME	"sqlite_sync_contacts_db"

QString SyncContactsThread::DBPath;

SyncContactsThread::SyncContactsThread(QObject *parent)
    : QThread(parent)
{
    moveToThread(this);

    DBPath =  g_AppSettingsFolderPath + "/outcall.db";
}

SyncContactsThread::~SyncContactsThread()
{

}

void SyncContactsThread::run()
{
    m_bStop = false;

    QMutex *mutex = &(g_pContactManager->m_LoadContactsMutex);

    if (mutex->tryLock(5000)==false)
    { // wait 5 seconds max.
        emit loadingContactsFinished(false, "Mutex lock failed.");
        return;
    }

    {
        // we remove connection at the end, so finish with QSqlDatabase object in this scope
        QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", SQLITE_DB_CONNECTION_NAME);
        db.setDatabaseName(DBPath);
        db.open();
    }

    QString resultErrorMessage;
    bool bContactsLoadedSuccessfully = true;

    if (global::IsOutlookInstalled()) {
        bContactsLoadedSuccessfully = false;

        QString db_path= g_pContactManager->getDBPath();
        QString log_path = g_AppSettingsFolderPath + "/outcall.log";

        //if outlook x64
        if (true/*IsOutlook64bit()*/) // always use helper app to avoid potential crashing in outcall
        {
            bool b64 = global::IsOutlook64bit();
            QString helper_path;
            if (b64)
                helper_path = g_AppDirPath + "/Outlook/x64/outlook_helper_x64.exe";
            else
                helper_path = g_AppDirPath + "/Outlook/outlook_helper.exe";

            QStringList arguments;
            QString exchange = "true";
            arguments << QString("LoadContacts").toLatin1().data() << db_path.toLatin1().data() << log_path.toLatin1().data() << exchange;

            //HelperProcess outlookHelperProcess(helper_path, arguments);
            QProcess outlookHelperProcess;
            if (b64)
                outlookHelperProcess.setWorkingDirectory(g_AppDirPath + "/Outlook/x64");
            else
                outlookHelperProcess.setWorkingDirectory(g_AppDirPath + "/Outlook");

            outlookHelperProcess.start(helper_path, arguments);
            if (!outlookHelperProcess.waitForStarted(10000)) {
                if (b64)
                    resultErrorMessage = "Outlook helper x64 could not be started.";
                else
                    resultErrorMessage = "Outlook helper could not be started.";
            } else {
                while (!m_bStop) {
                    if (outlookHelperProcess.waitForFinished(1000))
                        break;
                }
                if (m_bStop) {
                    outlookHelperProcess.kill();
                    resultErrorMessage = "Cancelled";
                } else {
                    resultErrorMessage = QString(outlookHelperProcess.readAllStandardOutput());

                    if (resultErrorMessage=="_ok_") {
                        bContactsLoadedSuccessfully = true;
                        resultErrorMessage = "";
                    }
                    else if (resultErrorMessage.isEmpty()) {
                        resultErrorMessage = tr("Unknown error. Check log files for more information.");
                    }

                    /*if (resultErrorMessage.isEmpty())
                        bContactsLoadedSuccessfully = true;*/
                }
            }
        }
    }

    QSqlDatabase::removeDatabase(SQLITE_DB_CONNECTION_NAME); // free resources
    mutex->unlock();

    emit loadingContactsFinished(bContactsLoadedSuccessfully, resultErrorMessage);
}


void SyncContactsThread::Stop()
{
    quit();
}
