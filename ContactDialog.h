#ifndef CONTACTDIALOG_H
#define CONTACTDIALOG_H

#include <QDialog>
#include <QHash>

class QTreeWidgetItem;

namespace Ui {
class ContactDialog;
}

class ContactDialog : public QDialog
{
    Q_OBJECT

public:
    ContactDialog(QWidget *parent = 0);
    ~ContactDialog();

    void setName(const QString &name);
    void setNumbers(QHash<QString, QString> &numbers);

protected:
    void clearCallTree();

protected slots:
    void onCallButton();
    void onCancelButton();
    void onSettingsChange();
    void onItemDoubleClicked(QTreeWidgetItem *item, int);

private:
    Ui::ContactDialog *ui;
};

#endif // CONTACTDIALOG_H
