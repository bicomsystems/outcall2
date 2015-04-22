#ifndef SPEEDDIALDIALOG_H
#define SPEEDDIALDIALOG_H

#include <QDialog>

namespace Ui {
class SpeedDialDialog;
}

class SpeedDialDialog : public QDialog
{
    Q_OBJECT

public:
    explicit SpeedDialDialog(QString code  = "", QString number = "", QWidget *parent = 0);
    ~SpeedDialDialog();

    QString getCode() const;
    QString getPhoneNumber() const;
    void removeCode(const QString &code);
    void updateSpeedDials(const QStringList &codes);

protected slots:
    void onAccept();

private:
    Ui::SpeedDialDialog *ui;

    static QStringList m_addedCodes;
    QString m_code;
};

#endif // SPEEDDIALDIALOG_H
