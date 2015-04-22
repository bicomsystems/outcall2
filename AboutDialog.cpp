#include "AboutDialog.h"
#include "ui_aboutdialog.h"
#include "QOutCallLabel.h"

#include <QDesktopServices>
#include <QUrl>

AboutDialog::AboutDialog(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::AboutDialog)
{
    ui->setupUi(this);
    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    connect(ui->pushButton, &QPushButton::clicked,   this, &QDialog::close);
    connect(ui->lblAvatar,  &QOutCallLabel::clicked, this, &AboutDialog::onClicked);
}

AboutDialog::~AboutDialog()
{
    delete ui;
}

void AboutDialog::onClicked()
{
    QString link = "http://www.bicomsystems.com";
    QDesktopServices::openUrl(QUrl(link));
}
