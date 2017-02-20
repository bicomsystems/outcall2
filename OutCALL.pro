#-------------------------------------------------
#
# Project created by QtCreator 2015-03-16T10:53:56
#
#-------------------------------------------------

QT       += core gui
QT       += network
QT       += sql

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = OutCALL
TEMPLATE = app


LIBS +=	Shell32.lib \
        Advapi32.lib \

TRANSLATIONS = translations/outcall_pt-br.ts \


SOURCES += main.cpp\
    DebugInfoDialog.cpp \
    AboutDialog.cpp \
    AddExtensionDialog.cpp \
    AddLanguageDialog.cpp \
    AsteriskManager.cpp \
    CallHistoryDialog.cpp \
    ContactDialog.cpp \
    ContactManager.cpp \
    Global.cpp \
    LocalServer.cpp \
    OutCALL.cpp \
    PlaceCallDialog.cpp \
    SearchBox.cpp \
    SpeedDialDialog.cpp \
    SyncContactsThread.cpp \
    SettingsDialog.cpp \
    PopupWindow.cpp \
    QOutCallLabel.cpp \
    Notifier.cpp

HEADERS  += \
    DebugInfoDialog.h \
    SyncContactsThread.h \
    AboutDialog.h \
    AddLanguageDialog.h \
    AsteriskManager.h \
    CallHistoryDialog.h \
    ContactDialog.h \
    ContactManager.h \
    Global.h \
    LocalServer.h \
    OutCALL.h \
    PlaceCallDialog.h \
    SearchBox.h \
    SpeedDialDialog.h \
    AddExtensionDialog.h \
    SettingsDialog.h \
    PopupWindow.h \
    QOutCallLabel.h \
    Notifier.h

FORMS    += \
    AboutDialog.ui \
    AddExtensionDialog.ui \
    AddLanguageDialog.ui \
    CallHistoryDialog.ui \
    ContactDialog.ui \
    DebugInfoDialog.ui \
    PlaceCallDialog.ui \
    SettingsDialog.ui \
    SpeedDialDialog.ui \
    PopupWindow.ui

RESOURCES += \
    images.qrc
