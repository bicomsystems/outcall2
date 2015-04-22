#ifndef PLACECALLDIALOG_H
#define PLACECALLDIALOG_H

#include <QDialog>

class Contact;
class QTreeWidgetItem;

namespace Ui {
class PlaceCallDialog;
}

class PlaceCallDialog : public QDialog
{
    Q_OBJECT

public:
    explicit PlaceCallDialog(QWidget *parent = 0);
    ~PlaceCallDialog();

    void show();

protected:
     void clearCallTree();

protected slots:
    void onCallButton();
    void onCancelButton();
    void onChangeContact(QString name);
    void onContactIndexChange(const QString &name);
    void onContactsLoaded(QList<Contact *> &contacts);
    void onSettingsChange();
    void onItemDoubleClicked(QTreeWidgetItem * item, int);
    void onItemClicked(QTreeWidgetItem * item, int);

private:
    Ui::PlaceCallDialog *ui;
    QList<Contact*> m_contacts;  /**< Contact list from Outlook */
};

#endif // PLACECALLDIALOG_H
