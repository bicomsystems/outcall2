#include "DebugInfoDialog.h"
#include "ui_debuginfodialog.h"
#include "Global.h"

DebugInfoDialog::DebugInfoDialog(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::DebugInfoDialog)
{
    ui->setupUi(this);

    ui->textEdit->setReadOnly(true);
    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    connect(ui->exit,  &QAbstractButton::clicked, this, &DebugInfoDialog::onExit);
    connect(ui->clear, &QAbstractButton::clicked, this, &DebugInfoDialog::onClear);
}

DebugInfoDialog::~DebugInfoDialog()
{
    delete ui;
}

void DebugInfoDialog::onClear() const
{
    ui->textEdit->clear();
}

void DebugInfoDialog::onExit()
{
    hide();
}

void DebugInfoDialog::updateDebug(const QString &info)
{
    ui->textEdit->appendPlainText(info);
}
