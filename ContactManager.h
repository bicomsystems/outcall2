#ifndef CONTACTMANAGER_H
#define CONTACTMANAGER_H

#include "ContactDialog.h"
#include "SyncContactsThread.h"

#include <QHash>
#include <QSqlDatabase>
#include <QTimer>
#include <QMutex>

class QString;
class SyncContactsThread;
class QTimer;
class QMutex;

class AddOutlookContactThread : public QThread {
    Q_OBJECT

public:

    AddOutlookContactThread(QString numbers, QString fullName = "");

    void stop();

signals:
    void addingContactThreadFinished(bool bOK, QString info);

protected:
    void run();

public slots:

private:

    HINSTANCE m_hOutlookDLLInstance;
    QString m_numbers;
    QString m_fullName;
    bool m_bStop;
};

class ViewOutlookContactThread : public QThread {
    Q_OBJECT

public:
    ViewOutlookContactThread(QString fullName, QString numbers = "");

    void stop();
signals:
    void viewingContactThreadFinished(bool bOK, QString info);

protected:
    void run();

private:
    HINSTANCE m_hOutlookDLLInstance;
    QString m_numbers;
    QString m_fullName;
    bool m_bStop;
};

class Contact {
public:
    Contact();
    void addNumber(const QString &type, const QString &number);

    QString name;
    QString formattedDescription;
    QHash<QString, QString> numbers;
};

class ContactManager : public QObject
{
    Q_OBJECT

public:
    explicit ContactManager(QObject *parent = 0);
    ~ContactManager();

    void openContactDialog(QString name, QHash<QString, QString> numbers);

    void activateDialog() const;

    QString getDBPath() const;

    QList<Contact*> getContacts();

    void refreshContacts();

    void stopSyncing();

    Contact* findContact(QString name);

    Contact* findContactByPhoneNumber(QString number);

    void startRefreshTimer();

    void addOutlookContact(QString numbers = "", QString fullName = "");

    void viewOutlookContact(QString fullName, QString numbers = "");

protected:
    void loadContacts();

    void createDBTable(QString table, QString query);

private slots:
    void loadingContactsThreadFinished(bool bOK, QString info);

    void loadContactsThreadStarted();

    void addingContactThreadFinished(bool bOK, QString info);
    void onRefreshContactsTimer();

signals:
    void contactsLoaded(QList<Contact*> &contacts);
    void syncing(bool status);

private:

    QSqlDatabase m_db;
    SyncContactsThread m_LoadContactsThread;
    QString DBPath;
    QHash<QString, Contact*> m_contacts; // name -> Contact
    QHash<QString, Contact*> m_PhoneNumbers; // number -> Contact (formatted phone numbers because of the performance)
    QTimer m_RefreshContactsTimer;
    ContactDialog *m_contactDialog;

    QMutex m_LoadContactsMutex;

    friend class SyncContactsThread;
};

extern ContactManager *g_pContactManager;

#endif // CONTACTMANAGER_H
