#include "CallHistoryDialog.h"
#include "ui_callhistorydialog.h"
#include "Global.h"
#include "ContactManager.h"

#include <QDebug>
#include <QList>
#include <QMessageBox>

CallHistoryDialog::CallHistoryDialog(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::CallHistoryDialog)
{
    ui->setupUi(this);

    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    connect(ui->closeButton,  &QPushButton::clicked, this, &QDialog::close);
    connect(ui->removeButton, &QPushButton::clicked, this, &CallHistoryDialog::onRemoveButton);
    connect(ui->callButton,   &QPushButton::clicked, this, &CallHistoryDialog::onCallClicked);
    connect(ui->addContactButton, &QPushButton::clicked, this, &CallHistoryDialog::onAddContact);

    ui->tabWidget->setCurrentIndex(0);

    // Load calls
    QVariantList missedCalls = global::getSettingsValue("missed", "calls").toList();
    for (int i = 0; i < missedCalls.size(); ++i)
    {
        QMap<QString, QVariant> call = missedCalls[i].toMap();
        addCall(call, MISSED);
    }

    QVariantList receivedCalls = global::getSettingsValue("received", "calls").toList();
    for (int i = 0; i < receivedCalls.size(); ++i)
    {
        QMap<QString, QVariant> call = receivedCalls[i].toMap();
        addCall(call, RECIEVED);
    }

    QVariantList placedCalls = global::getSettingsValue("placed", "calls").toList();
    for (int i = 0; i < placedCalls.size(); ++i)
    {
        QMap<QString, QVariant> call = placedCalls[i].toMap();
        addCall(call, PLACED);
    }
}

CallHistoryDialog::~CallHistoryDialog()
{
    delete ui;
}

void CallHistoryDialog::addCall(const QMap<QString, QVariant> &call, CallHistoryDialog::Calls calls)
{
    const QString from     = call.value("from").toString();
    const QString to       = call.value("to").toString();
    const QString protocol = call.value("protocol").toString();
    const QString dateTime = call.value("date_time").toString();
    QString callerIDName   = call.value("callerIDName").toString();

    QList<Contact*> contactList = g_pContactManager->getContacts();
    for(int i = 0; i < contactList.size(); ++i)
    {
        Contact *contact = contactList[i];
        QList<QString> numbers  = contact->numbers.values();
        if (numbers.contains(from))
        {
            callerIDName = contact->name;
            break;
        }
    }

    if (calls == MISSED)
    {
        QTreeWidgetItem *extensionItem = new QTreeWidgetItem(ui->treeWidgetMissed);
        if (callerIDName != "<unknown>")
            extensionItem->setText(0, callerIDName);

        extensionItem->setText(1, from);
        extensionItem->setText(2, to + "(" + protocol + ")");
        extensionItem->setText(3, dateTime);
    }
    else if (calls == RECIEVED)
    {
        QTreeWidgetItem *extensionItem = new QTreeWidgetItem(ui->treeWidgetReceived);
        if (callerIDName != "<unknown>")
            extensionItem->setText(0, callerIDName);

        extensionItem->setText(1, from);
        extensionItem->setText(2, to + "(" + protocol + ")");
        extensionItem->setText(3, dateTime);
    }
    else if (calls == PLACED)
    {
        QTreeWidgetItem *extensionItem = new QTreeWidgetItem(ui->treeWidgetPlaced);

        extensionItem->setText(0, from + "(" + protocol + ")");
        extensionItem->setText(1, to);
        extensionItem->setText(2, dateTime);
    }
}

void CallHistoryDialog::onCallClicked()
{
    QList<QTreeWidgetItem*> selectedItems;

    if (ui->tabWidget->currentIndex() == MISSED)
    {
       selectedItems = ui->treeWidgetMissed->selectedItems();

       if (selectedItems.size() == 0)
           return;

       QTreeWidgetItem *item = selectedItems.at(0);
       const QString from = item->text(1);
       int ind1 = item->text(2).indexOf("(");
       int ind2 = item->text(2).indexOf(")");
       const QString to = item->text(2).mid(0, ind1);
       const QString protocol = item->text(2).mid(ind1 + 1, ind2 - ind1 - 1);

       g_pAsteriskManager->originateCall(to, from, protocol, to);
    }
    else if (ui->tabWidget->currentIndex() == RECIEVED)
    {
       selectedItems = ui->treeWidgetReceived->selectedItems();

       if (selectedItems.size() == 0)
           return;

       QTreeWidgetItem *item = selectedItems.at(0);
       const QString from = item->text(1);
       int ind1 = item->text(2).indexOf("(");
       int ind2 = item->text(2).indexOf(")");
       const QString to = item->text(2).mid(0, ind1);
       const QString protocol = item->text(2).mid(ind1+1, ind2-ind1-1);

       g_pAsteriskManager->originateCall(to, from, protocol, to);
    }
    else if (ui->tabWidget->currentIndex() == PLACED)
    {
       selectedItems = ui->treeWidgetPlaced->selectedItems();

       if (selectedItems.size() == 0)
           return;

       QTreeWidgetItem *item = selectedItems.at(0);
       int ind1 = item->text(0).indexOf("(");
       int ind2 = item->text(0).indexOf(")");
       const QString from = item->text(0).mid(0, ind1);
       const QString to = item->text(1);
       const QString protocol = item->text(0).mid(ind1 + 1, ind2 - ind1 - 1);

       g_pAsteriskManager->originateCall(from, to, protocol, from);
    }
}

void CallHistoryDialog::onAddContact()
{
    if (ui->tabWidget->currentIndex() == MISSED)
    {
        QList<QTreeWidgetItem*> selectedItems = ui->treeWidgetMissed->selectedItems();
        if (selectedItems.size() == 0)
            return;

        QTreeWidgetItem *item = selectedItems.at(0);
        const QString name = item->text(0);
        const QString from = item->text(1);

        g_pContactManager->addOutlookContact(from, name);
    }
    else if (ui->tabWidget->currentIndex() == RECIEVED)
    {
        QList<QTreeWidgetItem*> selectedItems = ui->treeWidgetReceived->selectedItems();
        if (selectedItems.size() == 0)
            return;

        QTreeWidgetItem *item = selectedItems.at(0);
        const QString name = item->text(0);
        const QString from = item->text(1);

        g_pContactManager->addOutlookContact(from, name);
    }
    else if(ui->tabWidget->currentIndex() == PLACED)
    {
        QList<QTreeWidgetItem*> selectedItems = ui->treeWidgetPlaced->selectedItems();
        if (selectedItems.size() == 0)
            return;

        QTreeWidgetItem *item = selectedItems.at(0);
        int ind1 = item->text(1).indexOf("(");
        const QString to = item->text(1).mid(0, ind1);

        g_pContactManager->addOutlookContact(to, "");
    }
}

void CallHistoryDialog::onRemoveButton()
{
    auto messageBox = [=](bool openBox)->bool
    {
        if (!openBox)
            return false;

        QMessageBox::StandardButton reply;
        reply = QMessageBox::question(this, tr(APP_NAME), tr("Are you sure you want to delete the selected items ?"),
                                          QMessageBox::Yes|QMessageBox::No);

        if (reply == QMessageBox::No)
            return false;
        return true;
    };

    if (ui->tabWidget->currentIndex() == MISSED)
    {
        QList<QTreeWidgetItem*> selectedItems = ui->treeWidgetMissed->selectedItems();

        if (messageBox(selectedItems.size() > 0) == false)
            return;

        for (int i = 0; i < selectedItems.size(); ++i)
        {
            QTreeWidgetItem *item = selectedItems.at(i);
            int index = ui->treeWidgetMissed->indexOfTopLevelItem(item);
            ui->treeWidgetMissed->takeTopLevelItem(index);
        }
        global::removeSettinsKey("missed", "calls");

        QVariantList missedList;
        for (int i = 0; i < ui->treeWidgetMissed->topLevelItemCount(); ++i)
        {
            QTreeWidgetItem *item       = ui->treeWidgetMissed->topLevelItem(i);
            const QString callerIDName  = item->text(0);
            const QString from          = item->text(1);

            int ind1 = item->text(2).indexOf("(");
            int ind2 = item->text(2).indexOf(")");

            const QString to       = item->text(1).mid(0, ind1);
            const QString protocol = item->text(2).mid(ind1 + 1, ind2 - ind1 - 1);
            const QString dateTime = item->text(3);

            QMap<QString, QVariant> missed;
            missed.insert("from", from);
            missed.insert("to", to);
            missed.insert("protocol", protocol);
            missed.insert("date_time", dateTime);
            missed.insert("callerIDName", callerIDName);

            missedList.append(QVariant::fromValue(missed));
        }
        global::setSettingsValue("missed", missedList, "calls");

    }
    else if (ui->tabWidget->currentIndex() == RECIEVED)
    {
        QList<QTreeWidgetItem*> selectedItems = ui->treeWidgetReceived->selectedItems();

        if (messageBox(selectedItems.size() > 0) == false)
            return;

        for (int i = 0; i < selectedItems.size(); ++i)
        {
            QTreeWidgetItem *item = selectedItems.at(i);
            int index = ui->treeWidgetReceived->indexOfTopLevelItem(item);
            ui->treeWidgetReceived->takeTopLevelItem(index);
        }
        global::removeSettinsKey("received", "calls");

        QVariantList receivedList;

        for (int i = 0; i < ui->treeWidgetReceived->topLevelItemCount(); ++i)
        {
            QTreeWidgetItem *item       = ui->treeWidgetReceived->topLevelItem(i);
            const QString callerIDName  = item->text(0);
            const QString from          = item->text(1);

            int ind1 = item->text(2).indexOf("(");
            int ind2 = item->text(2).indexOf(")");

            const QString to       = item->text(1).mid(0, ind1);
            const QString protocol = item->text(2).mid(ind1 + 1, ind2 - ind1 - 1);
            const QString dateTime = item->text(3);

            QMap<QString, QVariant> received;
            received.insert("from", from);
            received.insert("to", to);
            received.insert("protocol", protocol);
            received.insert("date_time", dateTime);
            received.insert("callerIDName", callerIDName);

            receivedList.append(QVariant::fromValue(received));
        }
        global::setSettingsValue("received", receivedList, "calls");
    }
    else if (ui->tabWidget->currentIndex() == PLACED)
    {
        QList<QTreeWidgetItem*> selectedItems = ui->treeWidgetPlaced->selectedItems();

        if (messageBox(selectedItems.size() > 0) == false)
            return;

        for (int i = 0; i < selectedItems.size(); ++i)
        {
            QTreeWidgetItem *item = selectedItems.at(i);
            int index = ui->treeWidgetPlaced->indexOfTopLevelItem(item);
            ui->treeWidgetPlaced->takeTopLevelItem(index);
        }
        global::removeSettinsKey("placed", "calls");

        QVariantList placedList;

        for (int i = 0; i < ui->treeWidgetPlaced->topLevelItemCount(); ++i)
        {
            QTreeWidgetItem *item   = ui->treeWidgetPlaced->topLevelItem(i);
            const QString from      = item->text(0);

            int ind1 = item->text(1).indexOf("(");
            int ind2 = item->text(1).indexOf(")");

            const QString to        = item->text(1).mid(0, ind1);
            const QString protocol  = item->text(1).mid(ind1 + 1, ind2 - ind1 - 1);
            const QString dateTime  = item->text(2);

            QMap<QString, QVariant> placed;
            placed.insert("from", from);
            placed.insert("to", to);
            placed.insert("protocol", protocol);
            placed.insert("date_time", dateTime);

            placedList.append(QVariant::fromValue(placed));
        }
        global::setSettingsValue("placed", placedList, "calls");
    }
}
