#ifndef CALLHISTORYDIALOG_H
#define CALLHISTORYDIALOG_H

#include "AsteriskManager.h"

#include <QDialog>
#include <QMap>

namespace Ui {
class CallHistoryDialog;
}

class CallHistoryDialog : public QDialog
{
    Q_OBJECT

public:
    explicit CallHistoryDialog(QWidget *parent = 0);

    ~CallHistoryDialog();

    enum Calls
    {
        MISSED = 0,
        RECIEVED = 1,
        PLACED = 2
    };

    void addCall(const QMap<QString, QVariant> &call, Calls calls);

protected slots:

    void onAddContact();
    void onRemoveButton();
    void onCallClicked();

private:
    Ui::CallHistoryDialog *ui;
};

#endif // CALLHISTORYDIALOG_H
