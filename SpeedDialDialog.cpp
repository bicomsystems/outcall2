#include "SpeedDialDialog.h"
#include "ui_speeddialdialog.h"
#include "Global.h"

#include <QMessageBox>

QStringList SpeedDialDialog::m_addedCodes;

SpeedDialDialog::SpeedDialDialog(QString code, QString phoneNubmer, QWidget *parent) :
    QDialog(parent),
    ui(new Ui::SpeedDialDialog)
{
    ui->setupUi(this);

    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    QObject::connect(ui->buttonBox, &QDialogButtonBox::accepted, this, &SpeedDialDialog::onAccept);
    QObject::connect(ui->buttonBox, &QDialogButtonBox::rejected, this, &QDialog::reject);

    m_code = code;

    // Add code items.
    for (int i = 1; i < 100; ++i)
    {
        QString codes = QString::number(i);

        if (code.toInt() == i)
        {
            ui->codeBox->addItem(codes);
        }
        else if (m_addedCodes.contains(codes) == false)
        {
            ui->codeBox->addItem(codes);
        }
    }

    // Set current index.
    int index = ui->codeBox->findText(code);

    if (index == -1)
        index = 0;

    ui->codeBox->setCurrentIndex(index);
    ui->phoneEdit->setText(phoneNubmer);
}

SpeedDialDialog::~SpeedDialDialog()
{
    delete ui;
}

QString SpeedDialDialog::getCode() const
{
    return ui->codeBox->currentText();
}

QString SpeedDialDialog::getPhoneNumber() const
{
    return ui->phoneEdit->text();
}

void SpeedDialDialog::removeCode(const QString &code)
{
    m_addedCodes.removeOne(code);
}

void SpeedDialDialog::updateSpeedDials(const QStringList &codes)
{
    foreach(QString code, codes)
        m_addedCodes.append(code);
}

void SpeedDialDialog::onAccept()
{
    if (ui->phoneEdit->text().isEmpty())
    {
        MsgBoxWarning(tr("Invalid phone number!"));
    }
    else
    {
        QString currentCode = ui->codeBox->currentText();

        if (m_code.isEmpty() == false)
            m_addedCodes.removeOne(m_code);

        if (m_addedCodes.contains(currentCode) == false)
            m_addedCodes.append(currentCode);

        QDialog::accept();
    }
}
