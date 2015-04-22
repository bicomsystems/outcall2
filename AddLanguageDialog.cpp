#include "AddLanguageDialog.h"
#include "ui_addlanguagedialog.h"
#include "Global.h"

#include <QDir>
#include <QFile>
#include <QFileDialog>
#include <QPushButton>

AddLanguageDialog::AddLanguageDialog(QStringList countries, QWidget *parent) :
    QDialog(parent),
    ui(new Ui::AddLanguageDialog)
{
    ui->setupUi(this);

    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    connect(ui->addLanguageBtn, &QPushButton::clicked, this, &AddLanguageDialog::onAddLanguage);
    connect(ui->cancelBtn,      &QPushButton::clicked, this, &QDialog::close);

    QString country;
    ui->listWidgetLanguage->setSelectionMode(QAbstractItemView::SingleSelection);

    int pos;

    QListWidgetItem *item;

    for(int i = 0; i < countries.count(); ++i)
    {
        country = countries[i];

        pos = country.indexOf(" ");
        if (pos == -1)
            continue;

        item = new QListWidgetItem(country.mid(pos+1), ui->listWidgetLanguage);
        item->setData(Qt::UserRole, country.left(pos));
        ui->listWidgetLanguage->addItem(item);
    }
}

AddLanguageDialog::~AddLanguageDialog()
{
    delete ui;
}

void AddLanguageDialog::changeEvent(QEvent *e)
{
    QDialog::changeEvent(e);

    switch (e->type())
    {
    case QEvent::LanguageChange:
        ui->retranslateUi(this);
        break;
    default:
        break;
    }
}

void AddLanguageDialog::onAddLanguage()
{
    if (ui->listWidgetLanguage->selectedItems().count() == 0)
    {
        MsgBoxInformation(tr("Please select language to add first."));
        return;
    }

    QListWidgetItem *item = ui->listWidgetLanguage->selectedItems()[0];

    QString sourcePath = QFileDialog::getOpenFileName(this, tr("Open Language File"), g_AppDirPath, tr("Language Files (*.lang)"));
    QString destinationPath = g_LanguagesPath + "/" + item->data(Qt::UserRole).toString() + ".lang";

    if (!sourcePath.isEmpty())
    {
        if (QFile::exists(destinationPath))
        {
            if (!QFile::remove(destinationPath))
            {
                MsgBoxError(tr("Same language file already exists in the languages folder. It could not be removed. Perhaps you should try running application with administrative privileges."));
                return;
            }
        }
        if (!QFile::copy(sourcePath, destinationPath))
        {
            MsgBoxError(tr("Could not copy language file into languages folder. Perhaps you should try running application with administrative privileges."));
            return;
        }
        MsgBoxInformation(tr("Language file was added successfully."));
        accept();
    }
    else
    {
        return;
    }
}
