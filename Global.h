#ifndef GLOBAL_H
#define GLOBAL_H

#define LOG_ERROR				0
#define LOG_WARNING				1
#define LOG_INFORMATION			2
#define APP_NAME                "OutCALL"
#define ORGANIZATION_NAME       "Bicom Systems"

#include <QString>
#include <QVariant>
#include <QMessageBox>

extern QString g_LanguagesPath;
extern QString g_AppDirPath;
extern QString g_AppSettingsFolderPath;


QMessageBox::StandardButton MsgBoxInformation(const QString &text, QMessageBox::StandardButtons buttons = QMessageBox::Ok,
                                                     const QString &title=APP_NAME, QWidget *parent=0,
                                                     QMessageBox::StandardButton defaultButton = QMessageBox::NoButton);

QMessageBox::StandardButton MsgBoxError(const QString &text, QMessageBox::StandardButtons buttons = QMessageBox::Ok,
                                                     const QString &title=APP_NAME, QWidget *parent=0,
                                                     QMessageBox::StandardButton defaultButton = QMessageBox::NoButton);

QMessageBox::StandardButton MsgBoxWarning(const QString &text, QMessageBox::StandardButtons buttons = QMessageBox::Ok,
                                                     const QString &title=APP_NAME, QWidget *parent=0,
                                                     QMessageBox::StandardButton defaultButton = QMessageBox::NoButton);

namespace global {

    void setSettingsValue(const QString key, const QVariant value, const QString group = "");

    QVariant getSettingsValue(const QString key, const QString group = "", const QVariant value = QVariant());

    void removeSettinsKey(const QString key, const QString group = "");

    bool containsSettingsKey(const QString key, const QString group = "");

    QStringList getSettingKeys(const QString group);

    void log(const QString& msg, int type);

    void IntegrateIntoOutlook();

    bool IsOutlookInstalled();

    bool IsOutlook64bit(bool *bOutlookInstalled=NULL);

    bool EnableOutlookIntegration(bool bEnable);
}

#endif // GLOBAL_H
