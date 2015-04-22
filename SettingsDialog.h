#ifndef SETTINGSDIALOG_H
#define SETTINGSDIALOG_H

#include <QDialog>

class QAbstractButton;
class QTcpSocket;
class AddExtensionDialog;
class QTreeWidgetItem;

namespace Ui {
class SettingsDialog;
}

class SettingsDialog : public QDialog
{
    Q_OBJECT

public:
    explicit SettingsDialog(QWidget *parent = 0);
    ~SettingsDialog();

    void saveSettings();
    void loadSettings();
    void show();

protected:
    void handleButtonBox(QAbstractButton * button);
    void okPressed();
    void applyPressed();
    void onSpeedAddClicked();
    void onSpeedEditClicked();
    void onSpeedRemoveClicked();
    void onAddButtonClicked();
    void onRemoveButtonClicked();
    void onEditButtonClicked();
    void applySettings();
    void onReplaceDialAddClicked();
    void onReplaceDialRemoveClicked();
    void onReplaceDialTreeClicked(QTreeWidgetItem * item, int column);
    void onReplaceDialTreeDoubleClicked(QTreeWidgetItem * item, int column);
    void onMinCallerIDBoxChanged(int state);
    void onAddLanguageBtn();
    void loadLanguages();

private:
    Ui::SettingsDialog *ui;
    QTcpSocket *m_tcpSocket;
    AddExtensionDialog *m_addExtensionDialog;
    QStringList m_countries;
};

#endif // SETTINGSDIALOG_H
