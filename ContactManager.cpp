#include "ContactManager.h"
#include "ContactDialog.h"
#include "Global.h"

#include <QSqlQuery>
#include <QThread>
#include <QStringList>
#include <QDateTime>
#include <QDebug>
#include <QDir>
#include <QProcess>

ContactManager *g_pContactManager = nullptr;

Contact::Contact()
{
}

void Contact::addNumber(const QString &type, const QString &number)
{
    if (!numbers.values().contains(number))
        numbers.insertMulti(type, number);
}

AddOutlookContactThread::AddOutlookContactThread(QString numbers, QString fullName)
    : QThread()
{
    m_hOutlookDLLInstance = NULL;
    m_numbers = numbers;
    m_fullName = fullName;

    connect(this, SIGNAL(finished()), this, SLOT(deleteLater()));
}

void AddOutlookContactThread::run()
{
    QString result;
    bool bOK = false;

    QString log_path = g_AppSettingsFolderPath + "/outcall.log";

    //if outlook x64
    if (true/*IsOutlook64bit()*/)
    {
        bool b64 = global::IsOutlook64bit();
        QString helper_path;
        if (b64)
            helper_path = g_AppDirPath + "/Outlook/x64/outlook_helper_x64.exe";
        else
            helper_path = g_AppDirPath + "/Outlook/outlook_helper.exe";

        QStringList arguments;
        arguments << QString("AddNewContact").toLatin1().data() << log_path.toLatin1().data() << m_numbers.toLatin1().data() << "" << m_fullName.toLatin1().data();

        QProcess outlookHelperProcess;
        if (b64)
            outlookHelperProcess.setWorkingDirectory(g_AppDirPath + "/Outlook/x64");
        else
            outlookHelperProcess.setWorkingDirectory(g_AppDirPath + "/Outlook");

        outlookHelperProcess.start(helper_path, arguments);
        if (!outlookHelperProcess.waitForStarted(10000)) {
            if (b64)
                result = "Outlook helper x64 could not be started.";
            else
                result = "Outlook helper could not be started.";
        } else {
            outlookHelperProcess.waitForFinished(-1);
            result = QString(outlookHelperProcess.readAllStandardOutput());

            if (result=="_ok_") {
                bOK = true;
                result = "";
            }
            else if (result.isEmpty()) {
                result = tr("Unknown error. Check log files for more information.");
            }
        }
    }

    emit addingContactThreadFinished(bOK, result);
}

void AddOutlookContactThread::stop()
{
}

//*******************************************************************//
//*********************** ViewOutlookContactThreat*******************//
//*******************************************************************//

ViewOutlookContactThread::ViewOutlookContactThread(QString fullName, QString numbers)
    : QThread()
{
    m_hOutlookDLLInstance = NULL;
    m_numbers = numbers;
    m_fullName = fullName;

    connect(this, SIGNAL(finished()), this, SLOT(deleteLater()));
}

void ViewOutlookContactThread::run()
{
    QString result;
    bool bOK = false;

    QString log_path = g_AppSettingsFolderPath + "/outcall.log";

    bool b64 = global::IsOutlook64bit();
    QString helper_path;

    if (b64)
        helper_path = g_AppDirPath + "/Outlook/x64/outlook_helper_x64.exe";
    else
        helper_path = g_AppDirPath + "/Outlook/outlook_helper.exe";

    QStringList arguments;
    arguments << QString("ViewContact").toLatin1().data() << log_path.toLatin1().data() << m_numbers.toLatin1().data() << "" << m_fullName.toLatin1().data();

    QProcess outlookHelperProcess;
    if (b64)
        outlookHelperProcess.setWorkingDirectory(g_AppDirPath + "/Outlook/x64");
    else
        outlookHelperProcess.setWorkingDirectory(g_AppDirPath + "/Outlook");

    outlookHelperProcess.start(helper_path, arguments);
    if (!outlookHelperProcess.waitForStarted(10000))
    {
        if (b64)
            result = "Outlook helper x64 could not be started.";
        else
            result = "Outlook helper could not be started.";
    }
    else
    {
        outlookHelperProcess.waitForFinished(-1);
        result = QString(outlookHelperProcess.readAllStandardOutput());

        if (result=="_ok_")
        {
            bOK = true;
            result = "";
        }
        else if (result.isEmpty())
        {
            result = tr("Unknown error. Check log files for more information.");
        }
    }

    emit viewingContactThreadFinished(bOK, result);
}

void ViewOutlookContactThread::stop()
{
}

ContactManager::ContactManager(QObject *parent) : QObject(parent)
{
    g_pContactManager = this;

    m_contactDialog = new ContactDialog;

    DBPath =  g_AppSettingsFolderPath + "/outcall.db";

    m_db = QSqlDatabase::addDatabase("QSQLITE", "sqlite");
    m_db.setDatabaseName(DBPath);
    m_db.open();

    createDBTable("contacts", "CREATE TABLE contacts(name TEXT, field TEXT, value TEXT)");

    connect(&m_LoadContactsThread, SIGNAL(started()), this, SLOT(loadContactsThreadStarted()));
    connect(&m_LoadContactsThread, SIGNAL(loadingContactsFinished(bool, QString)), this, SLOT(loadingContactsThreadFinished(bool,QString)));

    loadContacts(); // load contacts from our db

    bool bContactsSynced = (global::getSettingsValue("LastContactSync", "", 2).toInt() != 2);

    connect(&m_RefreshContactsTimer, SIGNAL(timeout()), this, SLOT(onRefreshContactsTimer()));


    m_RefreshContactsTimer.setInterval(86400000); // 24 hours

    if (global::getSettingsValue("refresh_interval", "outlook").toBool())
        m_RefreshContactsTimer.start();

    if (!bContactsSynced)
        m_LoadContactsThread.start();
}

ContactManager::~ContactManager()
{
    m_LoadContactsThread.Stop();
    delete m_contactDialog;
}

void ContactManager::openContactDialog(QString name, QHash<QString, QString> numbers)
{ 
    m_contactDialog->setName(name);
    m_contactDialog->setNumbers(numbers);

    m_contactDialog->show();
    m_contactDialog->activateWindow();
}

void ContactManager::activateDialog() const
{
    m_contactDialog->activateWindow();
}

QString ContactManager::getDBPath() const
{
    return DBPath;
}

void ContactManager::createDBTable(QString table, QString query)
{
    QSqlQuery q(m_db);
    if (q.exec("select * from sqlite_master"))
    {
        while (q.next()) {
            if (q.value(1).toString().compare(table, Qt::CaseInsensitive)==0) {
                if (q.value(4).toString().compare(query, Qt::CaseInsensitive)!=0) {
                    q.exec("drop table " + table);
                    break;
                } else {
                    return;
                }
            }
        }
        q.exec(query);
    }
}

void ContactManager::loadContacts()
{
    if (m_LoadContactsMutex.tryLock(100)==false)
        return;

    if (m_db.isOpen())
    {
        Contact *pContact = NULL;
        qDeleteAll(m_contacts);
        m_contacts.clear();

        m_PhoneNumbers.clear();

        QSqlQuery q(m_db);

        if (q.exec("select name, field, value from contacts where field='numbers' order by name"))
        {
            QString name, field;
            QString previousName;
            QStringList numbers;
            QStringList number;
            QString phoneNumber;
            QString formattedDescription;

            while (q.next())
            {
                name = q.value(0).toString();
                field = q.value(1).toString();

                if (name!=previousName || pContact==NULL)
                {
                    pContact = new Contact;
                    pContact->name = name;
                    m_contacts.insertMulti(name, pContact);
                }

                if (field == "numbers")
                {
                    numbers = q.value(2).toString().split("|");
                    for (int i=0; i<numbers.count(); i++)
                    {
                        number = numbers[i].split("=");

                        if (number.count()==2)
                        {
                            phoneNumber = number[1];

                            pContact->addNumber(number[0], phoneNumber);

                            phoneNumber.replace("+", "00");
                            phoneNumber.replace(" ", "");
                            phoneNumber.replace("(", "");
                            phoneNumber.replace(")", "");
                            phoneNumber.replace("-", "");
                            phoneNumber.replace("/", "");
                            m_PhoneNumbers.insert(phoneNumber, pContact);
                        }
                    }
                }
                previousName = name;
            }

            QList<Contact*> contacts = m_contacts.values();
            QHash<QString, QString>::iterator iter;
            int cnt;

            for (int i=0; i<contacts.count(); i++)
            {
                pContact = contacts[i];
                formattedDescription = "";

                if (pContact->numbers.count()==0)
                {
                    formattedDescription = tr("No phone numbers");
                }
                else
                {
                    cnt=0;

                    for (iter=pContact->numbers.begin(); iter!=pContact->numbers.end() && cnt<1; iter++)
                    {
                        if (formattedDescription.isEmpty())
                            formattedDescription = QString("%1: %2").arg(iter.key()).arg(iter.value());
                        else
                            formattedDescription.append(QString(", %1: %2").arg(iter.key()).arg(iter.value()));
                        cnt++;
                    }

                    if (cnt<pContact->numbers.count())
                    {
                        formattedDescription.append(" ...");
                    }
                }
                pContact->formattedDescription = formattedDescription;
            }
        }
    }

    m_LoadContactsMutex.unlock();

    emit contactsLoaded(m_contacts.values());
}

void ContactManager::loadContactsThreadStarted()
{
    emit syncing(true);
}

void ContactManager::loadingContactsThreadFinished(bool bOK, QString info)
{
    if (bOK)
    {
        global::setSettingsValue("LastContactSyncInfo", tr("Last successfull sync: %1").arg(QDateTime::currentDateTime().toString("hh:mm ddd MM/dd/yy")));
        global::setSettingsValue("LastContactSync", 1);
        global::log("Contacts synced successfully.", LOG_INFORMATION);
    }
    else
    {
        global::setSettingsValue("LastContactSyncInfo", tr("%1 (on %2)").
                         arg(info).
                         arg(QDateTime::currentDateTime().toString("hh:mm ddd MM/dd/yy")));

        global::setSettingsValue("LastContactSync", 0);

        global::log(QString("======= Failed to sync contacts: %1 =======").arg(info), LOG_ERROR);
    }

    loadContacts();
    emit syncing(false);
}

QList<Contact*> ContactManager::getContacts()
{
    return m_contacts.values();
}

Contact* ContactManager::findContact(QString name)
{
    QHash<QString, Contact*>::iterator iter = m_contacts.find(name);
    if (iter != m_contacts.end())
        return iter.value();
    else
        return NULL;
}

void ContactManager::onRefreshContactsTimer()
{
   // m_RefreshContactsTimer.stop();
    m_LoadContactsThread.start();
}

void ContactManager::refreshContacts()
{
    m_RefreshContactsTimer.stop();
    m_LoadContactsThread.start();
}

void ContactManager::startRefreshTimer()
{
    m_RefreshContactsTimer.start();
}

void ContactManager::addOutlookContact(QString numbers, QString fullName)
{
    if (numbers.indexOf("=") == -1)
        numbers = "Business=" + numbers;

    AddOutlookContactThread *addContactThread = new AddOutlookContactThread(numbers, fullName);
    connect(addContactThread, SIGNAL(addingContactThreadFinished(bool, QString)),
            this, SLOT(addingContactThreadFinished(bool, QString)));
    addContactThread->start();
}

void ContactManager::viewOutlookContact(QString fullName, QString numbers)
{
    ViewOutlookContactThread *viewContactThread = new ViewOutlookContactThread(fullName, numbers);
    viewContactThread->start();
}

void ContactManager::addingContactThreadFinished(bool bOK, QString info)
{
    if (bOK)
    {
        refreshContacts();
    }
    else
    {
        MsgBoxError(tr("Failed to add Outlook contact. %1").arg(info));
    }
}
