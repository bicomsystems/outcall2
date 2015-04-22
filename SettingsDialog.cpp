#include "SettingsDialog.h"
#include "ui_settingsdialog.h"
#include "AddExtensionDialog.h"
#include "Global.h"
#include "ContactManager.h"
#include "SpeedDialDialog.h"
#include "AsteriskManager.h"
#include "AddLanguageDialog.h"
#include "Notifier.h"

#include <QAbstractButton>
#include <QAbstractSocket>
#include <QSettings>
#include <QKeyEvent>
#include <QMessageBox>
#include <QDir>

SettingsDialog::SettingsDialog(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::SettingsDialog),
    m_addExtensionDialog(nullptr)
{
    ui->setupUi(this);

    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    connect(ui->buttonBox, &QDialogButtonBox::clicked, this, &SettingsDialog::handleButtonBox);

    // Dialing rules
    connect(ui->speedAddButton,          &QPushButton::clicked,           this, &SettingsDialog::onSpeedAddClicked);
    connect(ui->speedEditButton,         &QPushButton::clicked,           this, &SettingsDialog::onSpeedEditClicked);
    connect(ui->speedRemoveButton,       &QPushButton::clicked,           this, &SettingsDialog::onSpeedRemoveClicked);
    connect(ui->replaceDialAddButton,    &QPushButton::clicked,           this, &SettingsDialog::onReplaceDialAddClicked);
    connect(ui->replaceDialRemoveButton, &QPushButton::clicked,           this, &SettingsDialog::onReplaceDialRemoveClicked);
    connect(ui->replaceTree,             &QTreeWidget::itemClicked,       this, &SettingsDialog::onReplaceDialTreeClicked);
    connect(ui->replaceTree,             &QTreeWidget::itemDoubleClicked, this, &SettingsDialog::onReplaceDialTreeDoubleClicked);

    // Extensions
    connect(ui->addButton,    &QPushButton::clicked, this, &SettingsDialog::onAddButtonClicked);
    connect(ui->removeButton, &QPushButton::clicked, this, &SettingsDialog::onRemoveButtonClicked);
    connect(ui->editButton,   &QPushButton::clicked, this, &SettingsDialog::onEditButtonClicked);
    ui->port->setValidator(new QIntValidator(0, 65535, this));

    // General
    for (int i = 2; i < 26; ++i)
        ui->callerIDList->addItem(tr("%1").arg(i));

    connect(ui->callerIDBox, &QCheckBox::stateChanged, this, &SettingsDialog::onMinCallerIDBoxChanged);
    connect(ui->languageButton, &QPushButton::clicked, this, &SettingsDialog::onAddLanguageBtn);

    QFile inputFile(g_LanguagesPath + "/languages.txt");
    inputFile.open(QIODevice::ReadOnly);
    QTextStream in(&inputFile);
    m_countries = in.readAll().split(QRegExp("(\\r\\n)|(\\n\\r)|\\r|\\n"), QString::SkipEmptyParts);
    ui->tabWidget->setCurrentIndex(0);

    loadLanguages();
    loadSettings();
    global::IntegrateIntoOutlook();
}

SettingsDialog::~SettingsDialog()
{
    delete ui;
}

void SettingsDialog::saveSettings()
{
    // General SettingsDialog
    global::setSettingsValue("auto_sign_in",  ui->autoSignIn->isChecked(),   "general");
    global::setSettingsValue("auto_startup",  ui->autoStartBox->isChecked(), "general");
    global::setSettingsValue("min_caller_state", ui->callerIDBox->isChecked(),  "general");
    global::setSettingsValue("min_caller_id", ui->callerIDList->currentText(), "general");
    global::setSettingsValue("language", ui->languageList->itemData(ui->languageList->currentIndex(), Qt::UserRole).toString(), "general");

    // Save Extension SettingsDialog
    global::setSettingsValue("servername", ui->serverName->text(), "settings");
    global::setSettingsValue("username",   ui->userName->text(),   "settings");
    QByteArray ba;
    ba.append(ui->password->text());
    global::setSettingsValue("password", ba.toBase64(),            "settings");
    global::setSettingsValue("port", ui->port->text().toUInt(),    "settings");

    // Save outlook settings
    global::setSettingsValue("outlook_integration", ui->enableOutlookBox->isChecked(), "outlook");
    global::setSettingsValue("refresh_interval", ui->refreshBox->isChecked(), "outlook");
    global::setSettingsValue("contact_inbound", ui->displayContactCallBox->isChecked(), "outlook");
    global::setSettingsValue("contact_unknown", ui->displayContactUnknownBox->isChecked(), "outlook");

    // Save Dialing Rules - Speed Dial
    global::removeSettinsKey("speed_dial");
    int speedDialNum = ui->speedTree->topLevelItemCount();
    for (int i = 0; i < speedDialNum; ++i)
    {
        QTreeWidgetItem *item = ui->speedTree->topLevelItem(i);
        const QString code = item->text(0);
        const QString phoneNumber = item->text(1);
        global::setSettingsValue(code, phoneNumber, "speed_dial");
    }

    // Save Dialing Rules - Replacement rules
    int nReplaceRulesList = ui->replaceTree->topLevelItemCount();
    QVariantList list;
    for (int i = 0; i < nReplaceRulesList; ++i)
    {
        QTreeWidgetItem *item = ui->replaceTree->topLevelItem(i);

        const QString text = item->text(0);
        const QString replacement    = item->text(1);
        int isRegEx           = (int)item->checkState(2);

        QMap<QString, QVariant> rules;
        rules.insert("replacement", replacement);
        rules.insert("text", text);
        rules.insert("regex", isRegEx);

        list.append(QVariant::fromValue(rules));
    }
    global::setSettingsValue("replacement_rules", list, "dial_rules");

    // Save Dialing Rules - General
    global::setSettingsValue("context", ui->contextLine->text(), "dial_rules");
    global::setSettingsValue("international", ui->internationalLine->text(), "dial_rules");
    global::setSettingsValue("prefix", ui->prefixLine->text(), "dial_rules");

    // Save extensions
    global::removeSettinsKey("extensions");
    int nRow = ui->treeWidget->topLevelItemCount();
    for (int i = 0; i < nRow; ++i)
    {
        QTreeWidgetItem *item = ui->treeWidget->topLevelItem(i);
        QString extension = item->text(0);
        QString protocol = item->text(1);
        global::setSettingsValue(extension, protocol, "extensions");
    }
}

void SettingsDialog::loadSettings()
{
    ui->serverName->setText(global::getSettingsValue("servername", "settings").toString());
    ui->userName->setText(global::getSettingsValue("username", "settings").toString());
    QByteArray password((global::getSettingsValue("password", "settings").toByteArray()));
    QString ba(QByteArray::fromBase64(password));
    ui->password->setText(ba);
    ui->port->setText(global::getSettingsValue("port", "settings", "5038").toString());

    // Load outlook settings
    ui->enableOutlookBox->setChecked(global::getSettingsValue("outlook_integration", "outlook", true).toBool());
    ui->refreshBox->setChecked(global::getSettingsValue("refresh_interval", "outlook", true).toBool());
    ui->displayContactCallBox->setChecked(global::getSettingsValue("contact_inbound", "outlook", false).toBool());
    ui->displayContactUnknownBox->setChecked(global::getSettingsValue("contact_unknown", "outlook", false).toBool());

    // Load General SettingsDialog
    ui->autoStartBox->setChecked(global::getSettingsValue("auto_startup", "general", true).toBool());
    ui->languageList->setCurrentText(global::getSettingsValue("language", "general").toString());

    bool autoSignIn = global::getSettingsValue("auto_sign_in",   "general", true).toBool();
    ui->autoSignIn->setChecked(autoSignIn);
    g_pAsteriskManager->setAutoSignIn(autoSignIn);

    bool minCallerId = global::getSettingsValue("min_caller_state", "general", false).toBool();
    ui->callerIDBox->setChecked(minCallerId);

    ui->callerIDList->setCurrentText(global::getSettingsValue("min_caller_id", "general").toString());
    ui->callerIDList->setEnabled(minCallerId);

    // Load extensions
    QStringList extensions = global::getSettingKeys("extensions");
    int nRows              = extensions.size();
    for (int i = 0; i < nRows; ++i)
    {
        const QString extension = extensions.at(i);
        const QString protocol  = global::getSettingsValue(extension, "extensions").toString();

        QTreeWidgetItem *extensionItem = new QTreeWidgetItem(ui->treeWidget);
        extensionItem->setText(0, extension);
        extensionItem->setText(1, protocol);
    }

    // Load dialing rules - Speed dial
    QStringList speedDials = global::getSettingKeys("speed_dial");
    int nSpeedDial = speedDials.size();
    for (int i = 0; i < nSpeedDial; ++i)
    {
        const QString code = speedDials.at(i);
        const QString phoneNumber = global::getSettingsValue(code, "speed_dial").toString();

        QTreeWidgetItem *extensionItem = new QTreeWidgetItem(ui->speedTree);
        extensionItem->setText(0, code);
        extensionItem->setText(1, phoneNumber);
    }
    ui->speedTree->sortItems(0, Qt::AscendingOrder);
    SpeedDialDialog speedDialDialog;
    speedDialDialog.updateSpeedDials(speedDials);

    // Load Dialing Rules - Replacement rules
    QVariantList list = global::getSettingsValue("replacement_rules", "dial_rules", QVariantList()).toList();
    for (int i = 0; i < list.size(); ++i)
    {
        QTreeWidgetItem *item = new QTreeWidgetItem(ui->replaceTree);
        QMap<QString, QVariant> rules = list[i].toMap();

        const QString text      = rules.value("text").toString();
        const QString replacement   = rules.value("replacement").toString();
        int isRegEx             = rules.value("regex").toInt();

        item->setText(0, text);
        item->setText(1, replacement);
        item->setCheckState(2, (Qt::CheckState)isRegEx);
    }
    global::setSettingsValue("replacement_rules", list, "dial_rules");

    // Load dialing rules - General
    ui->contextLine->setText(global::getSettingsValue("context", "dial_rules", "default").toString());
    ui->internationalLine->setText(global::getSettingsValue("international", "dial_rules").toString());
    ui->prefixLine->setText(global::getSettingsValue("prefix", "dial_rules").toString());
}

void SettingsDialog::show()
{
    ui->tabWidget->setCurrentIndex(0);
    ui->dialRulesTabs->setCurrentIndex(0);
    QDialog::show();
}

void SettingsDialog::handleButtonBox(QAbstractButton *button)
{
    if(ui->buttonBox->standardButton(button) == QDialogButtonBox::Ok)
    {
        okPressed();
    }
    else if (ui->buttonBox->standardButton(button) == QDialogButtonBox::Apply)
    {
        applyPressed();
    }
    if(ui->buttonBox->standardButton(button) == QDialogButtonBox::Cancel)
    {
        QDialog::close();
    }
}

void SettingsDialog::okPressed()
{
    saveSettings();
    applySettings();
    hide();
}

void SettingsDialog::applyPressed()
{
    saveSettings();
    applySettings();
}

void SettingsDialog::applySettings()
{
    global::IntegrateIntoOutlook();

    if (global::getSettingsValue("refresh_interval", "outlook").toBool())
        g_pContactManager->startRefreshTimer();

    QSettings startup("Microsoft", "Windows");
    if(ui->autoStartBox->isChecked())
    {
        QString path = g_AppDirPath;
        startup.setValue("/CurrentVersion/Run/OutCALL", path.replace("/", "\\") + QString("\\%1.exe").arg(APP_NAME));
    }
    else
    {
        startup.remove("/CurrentVersion/Run/OutCALL");
    }

    g_pAsteriskManager->setAutoSignIn(global::getSettingsValue("auto_sign_in", "general", true).toBool());
    g_Notifier->emitSettingsChanged();
}

/********************************************/
/****************General*********************/
/********************************************/

void SettingsDialog::onAddLanguageBtn()
{
    AddLanguageDialog dlg(m_countries, this);
    if (dlg.exec() == QDialog::Accepted)
    {
        loadLanguages();
    }
}

void SettingsDialog::loadLanguages()
{
    ui->languageList->clear();
    ui->languageList->addItem(tr("English (default)"), "");

    QStringList nameFilters;
    nameFilters.insert(0,"*.lang");
    QDir langDirectory(g_LanguagesPath);
    QStringList languages = langDirectory.entryList(nameFilters,QDir::Files,QDir::NoSort);
    QString country, file;

    int pos;

    for (int j = 0; j < languages.count(); ++j)
    {
        file = languages[j];
        file.replace(".lang", "");
        for (int i = 0; i < m_countries.count(); ++i)
        {
            country = m_countries[i];
            pos = country.indexOf(" ");
            if (pos==-1)
                continue;
            if (country.left(pos).compare(file, Qt::CaseInsensitive)==0)
            {
                ui->languageList->addItem(country.mid(pos+1), country.left(pos));
                break;
            }
        }
    }

    QString lang = global::getSettingsValue("language", "general").toString();
    if(lang == "")
        ui->languageList->setCurrentIndex(0);
    else
        ui->languageList->setCurrentIndex(ui->languageList->findData(lang, Qt::UserRole, Qt::MatchExactly));
}

/********************************************/
/****************Extensions******************/
/********************************************/

void SettingsDialog::onAddButtonClicked()
{
    AddExtensionDialog addExtensionDialog;
    addExtensionDialog.setWindowTitle("Add Extension");
    if(addExtensionDialog.exec())
    {
        QString extension = addExtensionDialog.getExtension();
        QString protocol = addExtensionDialog.getProtocol();

        QTreeWidgetItem *extensionItem = new QTreeWidgetItem();
        extensionItem->setText(0, extension);
        extensionItem->setText(1, protocol);
        extensionItem->setData(0, Qt::CheckStateRole, QVariant());

        ui->treeWidget->addTopLevelItem(extensionItem);
    }
}

void SettingsDialog::onRemoveButtonClicked()
{
    QList<QTreeWidgetItem*> selectedItems = ui->treeWidget->selectedItems();

    if (selectedItems.size() > 0)
    {
        QMessageBox::StandardButton reply;
        reply = QMessageBox::question(this, tr(APP_NAME),
                                      tr("Are you sure you want to delete the selected items ?"),
                                      QMessageBox::Yes|QMessageBox::No);

        if (reply == QMessageBox::No)
            return;

        for (int i = 0; i < selectedItems.size(); ++i)
        {
            int index = ui->treeWidget->indexOfTopLevelItem(selectedItems.at(i));
            ui->treeWidget->takeTopLevelItem(index);
        }
    }
}

void SettingsDialog::onEditButtonClicked()
{
    AddExtensionDialog editExtensionDialog;
    editExtensionDialog.setWindowTitle("Edit Extension");
    QList<QTreeWidgetItem*> selectedItems = ui->treeWidget->selectedItems();

    if (selectedItems.size())
    {
        QTreeWidgetItem *item = selectedItems.at(0);
        const QString extension = item->text(0);
        const QString protocol = item->text(1);

        editExtensionDialog.setExtension(extension);
        editExtensionDialog.setProtocol(protocol);

        if(editExtensionDialog.exec())
        {
            const QString newExtension = editExtensionDialog.getExtension();
            const QString newProtocol = editExtensionDialog.getProtocol();

            item->setText(0, newExtension);
            item->setText(1, newProtocol);
        }
    }
}

//****************************************************//
//**********************Dialing rules*****************//
//****************************************************//

void SettingsDialog::onSpeedAddClicked()
{
    SpeedDialDialog speedDialDialog;

    if(speedDialDialog.exec())
    {
        const QString code = speedDialDialog.getCode();
        const QString phoneNumber = speedDialDialog.getPhoneNumber();

        QTreeWidgetItem *item = new QTreeWidgetItem(ui->speedTree);
        item->setText(0, code);
        item->setText(1, phoneNumber);

        ui->speedTree->sortItems(0, Qt::AscendingOrder);
    }
}

void SettingsDialog::onSpeedEditClicked()
{
    QList<QTreeWidgetItem*> selectedItems = ui->speedTree->selectedItems();

    if (selectedItems.size())
    {
        QTreeWidgetItem *item = selectedItems.at(0);
        const QString code = item->text(0);
        const QString phoneNumber = item->text(1);

        SpeedDialDialog speedDialDialog(code, phoneNumber);

        if(speedDialDialog.exec())
        {
            const QString newPhoneNumber = speedDialDialog.getPhoneNumber();
            const QString newCode = speedDialDialog.getCode();

            item->setText(0, newCode);
            item->setText(1, newPhoneNumber);
        }
        ui->speedTree->sortItems(0, Qt::AscendingOrder);
    }
}

void SettingsDialog::onSpeedRemoveClicked()
{
    QList<QTreeWidgetItem*> selectedItems = ui->speedTree->selectedItems();

    if (selectedItems.size())
    {
        QMessageBox::StandardButton reply;
        reply = QMessageBox::question(this, tr(APP_NAME),
                                      tr("Are you sure you want to delete the selected items ?"),
                                      QMessageBox::Yes|QMessageBox::No);

        if (reply == QMessageBox::No)
              return;

        const QString code = selectedItems.at(0)->text(0);
        int index = ui->speedTree->indexOfTopLevelItem(selectedItems.at(0));
        ui->speedTree->takeTopLevelItem(index);
        SpeedDialDialog speedDialDialog;
        speedDialDialog.removeCode(code);
    }
}

void SettingsDialog::onReplaceDialAddClicked()
{
    QTreeWidgetItem *item = new QTreeWidgetItem(ui->replaceTree);
    item->setCheckState(2, Qt::Unchecked);
}

void SettingsDialog::onReplaceDialRemoveClicked()
{
    QList<QTreeWidgetItem*> selectedItems = ui->replaceTree->selectedItems();

    if (selectedItems.size())
    {
        int index = ui->replaceTree->indexOfTopLevelItem(selectedItems.at(0));
        ui->replaceTree->takeTopLevelItem(index);
    }
}

void SettingsDialog::onReplaceDialTreeClicked(QTreeWidgetItem * item, int column)
{
    if (column == 2)
    {
        item->setFlags(Qt::ItemIsEnabled | Qt::ItemIsUserCheckable | Qt::ItemIsSelectable);
    }
}

void SettingsDialog::onReplaceDialTreeDoubleClicked(QTreeWidgetItem * item, int column)
{
    if (column == 0 || column == 1)
    {
        item->setFlags(Qt::ItemIsEnabled | Qt::ItemIsEditable | Qt::ItemIsSelectable);
    }
}

void SettingsDialog::onMinCallerIDBoxChanged(int state)
{
    ui->callerIDList->setEnabled(state);
}
