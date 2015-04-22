#include "Global.h"

#include <Windows.h>
#include <TCHAR.H>
#include <shlobj.h>

#include <QSettings>
#include <QFile>
#include <QDir>
#include <QDateTime>
#include <QDebug>
#include <QApplication>

extern QString g_LanguagesPath = "";
extern QString g_AppDirPath = "";
extern QString g_AppSettingsFolderPath = "";

QMessageBox::StandardButton MsgBoxInformation(const QString &text, QMessageBox::StandardButtons buttons, const QString &title, QWidget *parent,
        QMessageBox::StandardButton defaultButton) {
        return QMessageBox::information(parent, title, text, buttons, defaultButton);
}

QMessageBox::StandardButton MsgBoxError(const QString &text, QMessageBox::StandardButtons buttons, const QString &title, QWidget *parent,
         QMessageBox::StandardButton defaultButton) {
         return QMessageBox::critical(parent, title, text, buttons, defaultButton);
}

QMessageBox::StandardButton MsgBoxWarning(const QString &text, QMessageBox::StandardButtons buttons, const QString &title, QWidget *parent,
         QMessageBox::StandardButton defaultButton) {
         return QMessageBox::warning(parent, title, text, buttons, defaultButton);
}

void global::setSettingsValue(QString key, QVariant value, QString group)
{
    QSettings settings(ORGANIZATION_NAME, APP_NAME);
    if (!group.isEmpty())
        settings.beginGroup(group);
    settings.setValue(key, value);
}

QVariant global::getSettingsValue(const QString key, const QString group, const QVariant defaultValue)
{
    QSettings settings(ORGANIZATION_NAME, APP_NAME);
    if (!group.isEmpty())
        settings.beginGroup(group);
    return settings.value(key, defaultValue);
}

void global::removeSettinsKey(const QString key, const QString group)
{
    QSettings settings(ORGANIZATION_NAME, APP_NAME);
    if (!group.isEmpty())
        settings.beginGroup(group);
    settings.remove(key);
}

bool global::containsSettingsKey(const QString key, const QString group)
{
    if (key.isEmpty())
        return false;

    QSettings settings(ORGANIZATION_NAME, APP_NAME);
    if (!group.isEmpty())
        settings.beginGroup(group);
    return settings.contains(key);
}

QStringList global::getSettingKeys(const QString group)
{
    QSettings settings(ORGANIZATION_NAME, APP_NAME);
    if (!group.isEmpty())
        settings.beginGroup(group);
    return settings.childKeys();
}

void global::log(const QString &msg, int type)
{
    static qint64 max_size = 1*1024*1024; // max 1 Megabyte
    QFile f (g_AppSettingsFolderPath + "/outcall.log");
    qint64 size = f.size();
    bool success;

    if (size > max_size)
        success = f.open(QIODevice::WriteOnly);
    else
        success = f.open(QIODevice::Append);

    if (success)
    {
        QString stype;
        switch (type)
        {
        case LOG_INFORMATION:
            stype = "Information";
            break;

        case LOG_WARNING:
            stype = "Warning";
            break;

        case LOG_ERROR:
            stype = "Error";
            break;

        default:
            break;
        }

        f.write(QString("%1 - %2: %3\r\n").arg(QDateTime::currentDateTime().toString("dd.MM hh:mm:ss")).arg(stype).arg(msg).toLatin1());
        f.close();
    }
}

void global::IntegrateIntoOutlook()
{
    if (IsOutlookInstalled())
        EnableOutlookIntegration(getSettingsValue("outlook_integration", "outlook").toBool());
    else
        log("Outlook is not installed.", LOG_INFORMATION);
}

bool global::IsOutlookInstalled()
{
    bool bOutlookInstalled = false;
    IsOutlook64bit(&bOutlookInstalled);
    return bOutlookInstalled;
}

bool global::IsOutlook64bit(bool *bOutlookInstalled)
{
    if (bOutlookInstalled)
        *bOutlookInstalled = false;

    bool bOutlook64bit = false;
    HKEY hKey;
    DWORD buffersize = 1024;
    TCHAR buffer[1024];

    if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\OUTLOOK.exe"), 0, KEY_READ, &hKey)==ERROR_SUCCESS)
    {
        if (RegQueryValueEx(hKey,TEXT("Path"),NULL,NULL,(LPBYTE) buffer, &buffersize)==ERROR_SUCCESS) {
            if (bOutlookInstalled && QFile::exists(QString(QString::fromUtf16((const ushort*)buffer) + "\\OUTLOOK.exe"))) {
                *bOutlookInstalled = true;
            }

            DWORD dwBinaryType;
            if (::GetBinaryType((LPCTSTR)QString(QString::fromUtf16((const ushort*)buffer) + "\\OUTLOOK.exe").utf16(), &dwBinaryType))
            {
                if (SCS_64BIT_BINARY == dwBinaryType)
                {
                    // Detected 64-bit Office
                    bOutlook64bit = true;
                }
            }
        }
        RegCloseKey(hKey);
    }
    else if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\OUTLOOK.exe"), 0, KEY_READ | KEY_WOW64_64KEY, &hKey)==ERROR_SUCCESS)
    {
        if (RegQueryValueEx(hKey,(LPCTSTR)_T("Path"),NULL,NULL,(LPBYTE) buffer, &buffersize)==ERROR_SUCCESS) {
            if (bOutlookInstalled && QFile::exists(QString(QString::fromUtf16((const ushort*)buffer) + "\\OUTLOOK.exe"))) {
                *bOutlookInstalled = true;
            }

            DWORD dwBinaryType;
            if (::GetBinaryType((LPCTSTR)QString(QString::fromUtf16((const ushort*)buffer) + "\\OUTLOOK.exe").utf16(), &dwBinaryType))
            {
                if (SCS_64BIT_BINARY == dwBinaryType)
                {
                    // Detected 64-bit Office
                    bOutlook64bit = true;
                }
            }
        }
        RegCloseKey(hKey);
    }
    return bOutlook64bit;
}

bool global::EnableOutlookIntegration(bool bEnable)
{
    HKEY hKey;

    bool bOutlook2010 = false;
    bool bOpen64bitRegistryView = false;

    // remove Outlook cache file (for some versions of Outlook this is required to load/unload plugin)
    TCHAR path[512];
    if (SUCCEEDED(SHGetFolderPath(NULL,CSIDL_LOCAL_APPDATA,NULL,0,path)))
    {
        try
        {
            QString cache_path = QString::fromUtf16((const ushort*)path) + "\\Microsoft\\Outlook\\extend.dat";
            QFile(cache_path).remove(); //this will refresh outlook cache for plugins (addins)
        }
        catch (...)
        {
        }
    }

    // NOTE about Windows 7 / Vista virtualization:
    // IF the app is not elevated and tries to write to HKLM\Software it actually writes to HKEY_USERS\<User SID>_Classes\VirtualStore\Machine\Software
    // we cannot really know if Reg API succeded or not

    hKey = NULL;
    for (int i=0; i<4; i++) {
        QString version;
        if (i==0)
            version = "14.0";
        else if (i==1)
            version = "15.0";
        else if (i==2)
            version = "16.0"; // future versions of outlook
        else if (i==3)
            version = "17.0"; // future versions of outlook

        QString key = QString("SOFTWARE\\Microsoft\\Office\\%1\\Common\\InstallRoot").arg(version);

        if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, (LPCTSTR)key.utf16(), 0, KEY_READ | KEY_WOW64_64KEY, &hKey)==ERROR_SUCCESS)
        {
            bOutlook2010 = true;
            break;
            //bOpen64bitRegistryView = true;
        }
        else if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, (LPCTSTR)key.utf16(), 0, KEY_READ, &hKey)==ERROR_SUCCESS) {
            bOutlook2010 = true;
            break;
        }
    }

    if (hKey)
    {
        RegCloseKey(hKey);
        hKey = NULL;
    }

    bOpen64bitRegistryView = IsOutlook64bit();

    if (bOutlook2010)
    {
        global::log(bOpen64bitRegistryView?"Detected Outlook 2010 (or greater) 64 bit.":"Detected Outlook 2010 (or greater) 32 bit.", LOG_INFORMATION);
    }
    else
    {
        global::log(bOpen64bitRegistryView?"Detected Outlook (lower than 2010) 64 bit.":"Detected Outlook (lower than 2010) 32 bit.", LOG_INFORMATION);
    }

    if (bOutlook2010)
    {
        QString plugin_key = "SOFTWARE\\Microsoft\\Office\\Outlook\\Addins\\OutCall.OutlookAddin";
        REGSAM access = bOpen64bitRegistryView?(KEY_READ | KEY_WRITE | KEY_WOW64_64KEY):(KEY_READ | KEY_WRITE);
        if (bEnable) {
            DWORD dwDisp = 0;
            LPDWORD lpdwDisp = &dwDisp;

            if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, (LPCTSTR)plugin_key.utf16(), REG_NONE, NULL, REG_OPTION_NON_VOLATILE, access, NULL, &hKey, lpdwDisp)==ERROR_SUCCESS)
            {
                QString pluginPath = QApplication::applicationDirPath().replace("/", "\\") + "\\Outlook\\OutlookPlugin.vsto|vstolocal";
                pluginPath.replace("/", "\\"); // because of '/' in the path, vstolocal was not working, Outlook would disable the Addin
                QString manifestPath = "file:///" + pluginPath;

                DWORD behavior = 3;
                QString name = "OutCALLPlugin";

                bool bSuccess = ((RegSetValueEx(hKey, TEXT("Description"), REG_NONE, REG_SZ, (LPBYTE)name.utf16(), name.length()*2)==ERROR_SUCCESS) &&
                                (RegSetValueEx(hKey, TEXT("FriendlyName"), REG_NONE, REG_SZ, (LPBYTE)name.utf16(), name.length()*2)==ERROR_SUCCESS) &&
                                (RegSetValueEx(hKey, TEXT("LoadBehavior"), REG_NONE, REG_DWORD,  (LPBYTE) &behavior, sizeof(DWORD))==ERROR_SUCCESS) &&
                                (RegSetValueEx(hKey, TEXT("Manifest"), REG_NONE, REG_SZ, (LPBYTE)manifestPath.utf16(), manifestPath.length()*2)==ERROR_SUCCESS));

                RegCloseKey(hKey);

                return bSuccess;
            }
            else
            {
                return false;
            }
        }
        else
        {
            if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Office\\Outlook\\Addins\\OutCall.OutlookAddin"), 0, access, &hKey)==ERROR_SUCCESS)
            {
                bool bSuccess = (RegDeleteValue(hKey, TEXT("Description"))==ERROR_SUCCESS &&
                                 RegDeleteValue(hKey, TEXT("FriendlyName"))==ERROR_SUCCESS &&
                                 RegDeleteValue(hKey, TEXT("LoadBehavior"))==ERROR_SUCCESS &&
                                 RegDeleteValue(hKey, TEXT("Manifest"))==ERROR_SUCCESS);
                /*DWORD behavior = 2;
                bool bSuccess = (RegSetValueEx(hKey, TEXT("LoadBehavior"), REG_NONE, REG_DWORD,  (LPBYTE) &behavior, sizeof(DWORD))==ERROR_SUCCESS);*/

                RegCloseKey(hKey);
                return bSuccess;
            }
            else
            {
                return false;
            }
        }
    }
    return false;
}




