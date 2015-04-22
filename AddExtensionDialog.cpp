#include "AddExtensionDialog.h"
#include "Global.h"
#include "ui_addextensiondialog.h"

AddExtensionDialog::AddExtensionDialog(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::AddExtensionDialog)
{
    ui->setupUi(this);

    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    QObject::connect(ui->buttonBox, &QDialogButtonBox::accepted, this, &AddExtensionDialog::onAccept);
    QObject::connect(ui->buttonBox, &QDialogButtonBox::rejected, this, &QDialog::reject);

    ui->extenLine->setValidator(new QIntValidator(this));
}

AddExtensionDialog::~AddExtensionDialog()
{
    delete ui;
}

QString AddExtensionDialog::getExtension() const
{
    return ui->extenLine->text();
}

QString AddExtensionDialog::getProtocol() const
{
    return ui->protocol->currentText();
}

void AddExtensionDialog::setExtension(const QString &extension)
{
    ui->extenLine->setText(extension);
}

void AddExtensionDialog::setProtocol(const QString &protocol)
{
    int index = ui->protocol->findText(protocol);
    ui->protocol->setCurrentIndex(index);
}

void AddExtensionDialog::onAccept()
{
    if (ui->extenLine->text().isEmpty())
        MsgBoxWarning(tr("Invalid phone number!"));
    else
        QDialog::accept();
}
