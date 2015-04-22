#ifndef ADDLANGUAGEDIALOG_H
#define ADDLANGUAGEDIALOG_H

#include <QDialog>

namespace Ui {
class AddLanguageDialog;
}

class AddLanguageDialog : public QDialog
{
    Q_OBJECT

public:
    AddLanguageDialog(QStringList countries, QWidget *parent = 0);
    ~AddLanguageDialog();

protected:
    void changeEvent(QEvent *e);

protected slots:
    void onAddLanguage();

private:
    Ui::AddLanguageDialog *ui;
};

#endif // ADDLANGUAGEDIALOG_H
