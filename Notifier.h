#ifndef SETTINGSNOTIFIER_H
#define SETTINGSNOTIFIER_H

#include<QObject>

class Notifier : public QObject
{
    Q_OBJECT

public:
    Notifier();
    ~Notifier();

    void emitSettingsChanged();

signals:
    void settingsChanged();
};

extern Notifier *g_Notifier;

#endif // SETTINGSNOTIFIER_H
