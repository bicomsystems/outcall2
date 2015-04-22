#include "ui_contactdialog.h"
#include "ContactDialog.h"
#include "AsteriskManager.h"
#include "Global.h"
#include "Notifier.h"

ContactDialog::ContactDialog(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::ContactDialog)
{
    ui->setupUi(this);

    connect(ui->callButton,     &QPushButton::clicked,              this, &ContactDialog::onCallButton);
    connect(ui->cancelButton,   &QPushButton::clicked,              this, &ContactDialog::onCancelButton);
    connect(g_Notifier,         &Notifier::settingsChanged,         this, &ContactDialog::onSettingsChange);
    connect(ui->treeWidget,     &QTreeWidget::itemDoubleClicked,    this, &ContactDialog::onItemDoubleClicked);

    QStringList extensions = global::getSettingKeys("extensions");
    for (int i = 0; i < extensions.size(); ++i)
        ui->extensionBox->addItem(extensions[i]);

    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);
}

ContactDialog::~ContactDialog()
{
    delete ui;
}

void ContactDialog::setName(const QString &name)
{
    ui->contactName->setText(name);
}

void ContactDialog::setNumbers(QHash<QString, QString> &numbers)
{
    clearCallTree();

    QList<QString> numList = numbers.keys();

    for(int i = 0; i < numList.size(); ++i)
    {
        QString type   = numList[i];
        QString number = numbers.value(type);

        QTreeWidgetItem *extensionItem = new QTreeWidgetItem(ui->treeWidget);
        extensionItem->setText(1, type);
        extensionItem->setText(0, number);
    }
}

void ContactDialog::onCallButton()
{
    QList<QTreeWidgetItem*> selectedItems = ui->treeWidget->selectedItems();

    if (selectedItems.size())
    {
        QTreeWidgetItem *item  = selectedItems.at(0);
        const QString number   = item->text(0);
        const QString from     = ui->extensionBox->currentText();
        const QString protocol = global::getSettingsValue(from, "extensions").toString();

        g_pAsteriskManager->originateCall(from, number, protocol, from);

        QDialog::close();
    }
}

void ContactDialog::onItemDoubleClicked(QTreeWidgetItem *item, int)
{
    QString number = item->text(0);
    const QString from     = ui->extensionBox->currentText();
    const QString protocol = global::getSettingsValue(from, "extensions").toString();

    g_pAsteriskManager->originateCall(from, number, protocol, from);

    QDialog::close();
}

void ContactDialog::onCancelButton()
{
    QDialog::close();
}

void ContactDialog::onSettingsChange()
{
    ui->extensionBox->clear();

    QStringList extensions = global::getSettingKeys("extensions");
    for (int i = 0; i < extensions.size(); ++i)
        ui->extensionBox->addItem(extensions[i]);
}

void ContactDialog::clearCallTree()
{
    int n = ui->treeWidget->topLevelItemCount();
    for (int i = 0; i < n; ++i)
        ui->treeWidget->takeTopLevelItem(i);

    if (n > 0)
        clearCallTree();
}
