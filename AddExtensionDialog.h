#ifndef ADDEXTENSIONDIALOG_H
#define ADDEXTENSIONDIALOG_H

#include <QDialog>

class QString;

namespace Ui {
class AddExtensionDialog;
}

class AddExtensionDialog : public QDialog
{
    Q_OBJECT

public:
    explicit AddExtensionDialog(QWidget *parent = 0);
    ~AddExtensionDialog();

    QString getExtension() const;
    QString getProtocol() const;

    void setExtension(const QString &extension);
    void setProtocol(const QString &protocol);

protected slots:
    void onAccept();

private:
    Ui::AddExtensionDialog *ui;
};

#endif // ADDEXTENSIONDIALOG_H
