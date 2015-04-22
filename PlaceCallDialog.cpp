#include "PlaceCallDialog.h"
#include "ui_placecalldialog.h"
#include "Global.h"
#include "ContactManager.h"
#include "SearchBox.h"
#include "AsteriskManager.h"
#include "Notifier.h"

#include <QHash>

PlaceCallDialog::PlaceCallDialog(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::PlaceCallDialog)
{
    ui->setupUi(this);

    connect(ui->callButton,    &QPushButton::clicked,           this, &PlaceCallDialog::onCallButton);
    connect(ui->cancelButton,  &QPushButton::clicked,           this, &PlaceCallDialog::onCancelButton);
    connect(ui->searchLine,    &SearchBox::selected,            this, &PlaceCallDialog::onChangeContact);
    connect(ui->treeWidget,    &QTreeWidget::itemClicked,       this, &PlaceCallDialog::onItemClicked);
    connect(ui->treeWidget,    &QTreeWidget::itemDoubleClicked, this, &PlaceCallDialog::onItemDoubleClicked);
    connect(g_pContactManager, &ContactManager::contactsLoaded, this, &PlaceCallDialog::onContactsLoaded);
    connect(g_Notifier,        &Notifier::settingsChanged,      this, &PlaceCallDialog::onSettingsChange);

    void (QComboBox:: *sig)(const QString&) = &QComboBox::currentIndexChanged;
    connect(ui->contactBox, sig, this, &PlaceCallDialog::onContactIndexChange);

    QStringList extensions = global::getSettingKeys("extensions");

    for (int i = 0; i < extensions.size(); ++i)
        ui->fromBox->addItem(extensions[i]);

    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    m_contacts = g_pContactManager->getContacts();
    QStringList data;

    foreach(Contact *contact, m_contacts)
    {
        data << contact->name;
        ui->contactBox->addItem(contact->name);
    }
    ui->searchLine->setData(data);
}

PlaceCallDialog::~PlaceCallDialog()
{
    delete ui;
}

void PlaceCallDialog::show()
{
    ui->contactBox->setCurrentIndex(0);
    ui->searchLine->clear();
    QDialog::show();
}

void PlaceCallDialog::onCallButton()
{
    if (!ui->phoneLine->text().isEmpty())
    {
        const QString number   = ui->phoneLine->text();
        const QString from     = ui->fromBox->currentText();
        const QString protocol = global::getSettingsValue(from, "extensions").toString();

        g_pAsteriskManager->originateCall(from, number, protocol, from);

        QDialog::close();
    }
}

void PlaceCallDialog::onCancelButton()
{
    QDialog::close();
}

void PlaceCallDialog::onChangeContact(QString name)
{
    Contact *contact = g_pContactManager->findContact(name);

    clearCallTree();

    if (contact != NULL)
    {
        QList<QString> numbers = contact->numbers.keys();

        QString number, type;

        for (int i = 0; i < numbers.size(); ++i)
        {
            type   = numbers[i];
            number = contact->numbers.value(type);

            QTreeWidgetItem *item = new QTreeWidgetItem(ui->treeWidget);
            item->setText(0, number);
            item->setText(1, type);
        }
    }
    QTreeWidgetItem *firstItem = ui->treeWidget->topLevelItem(0);
    if (firstItem)
    {
        firstItem->setSelected(true);
        QString number = firstItem->text(0);
        ui->phoneLine->setText(number);
    }
    else
    {
        ui->phoneLine->clear();
    }
}

void PlaceCallDialog::onItemDoubleClicked(QTreeWidgetItem *item, int)
{
    QString number = item->text(0);
    const QString from = ui->fromBox->currentText();
    const QString protocol = global::getSettingsValue(from, "extensions").toString();
    g_pAsteriskManager->originateCall(from, number, protocol, from);

    QDialog::close();
}

void PlaceCallDialog::onItemClicked(QTreeWidgetItem *item, int)
{
    QString number = item->text(0);
    ui->phoneLine->setText(number);
}

void PlaceCallDialog::onContactIndexChange(const QString &name)
{
    ui->searchLine->setText(name);
    onChangeContact(name);
}

void PlaceCallDialog::onContactsLoaded(QList<Contact *> &contacts)
{
    QStringList data;
    ui->contactBox->clear();

    foreach(Contact *contact, contacts)
    {
        data << contact->name;
        ui->contactBox->addItem(contact->name);
    }

    ui->searchLine->setData(data);
}

void PlaceCallDialog::clearCallTree()
{
    int n = ui->treeWidget->topLevelItemCount();
    for (int i = 0; i < n; ++i)
        ui->treeWidget->takeTopLevelItem(i);

    if (n > 0)
        clearCallTree();
}

void PlaceCallDialog::onSettingsChange()
{
    ui->fromBox->clear();

    QStringList extensions = global::getSettingKeys("extensions");

    for (int i = 0; i < extensions.size(); ++i)
        ui->fromBox->addItem(extensions[i]);
}
