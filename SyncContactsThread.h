#ifndef SYNCCONTACTSTHREAD_H
#define SYNCCONTACTSTHREAD_H

#include <windows.h>

#include <QThread>

class SyncContactsThread : public QThread
{
    Q_OBJECT

public:
    explicit SyncContactsThread(QObject *parent = 0);
    ~SyncContactsThread();

    static QString DBPath;

    void Stop();

protected:
    void run();

signals:
    void loadingContactsFinished(bool bOK, QString info);

private:
    HINSTANCE m_hOutlookDLLInstance;
    bool m_bStop;
};

#endif // SYNCCONTACTSTHREAD_H
