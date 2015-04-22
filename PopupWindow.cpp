#include "PopupWindow.h"
#include "ui_PopupWindow.h"

#include <QDesktopWidget>
#include <QMouseEvent>

QList<PopupWindow*> PopupWindow::m_PopupWindows;
int PopupWindow::m_nLastWindowPosition = 0;

#define TASKBAR_ON_TOP		1
#define TASKBAR_ON_LEFT		2
#define TASKBAR_ON_RIGHT	3
#define TASKBAR_ON_BOTTOM	4

#define TIME_TO_SHOW	700//1500 //msec
#define TIME_TO_LIVE	5000 //msec

PopupWindow::PopupWindow(const PWInformation& pwi, QWidget *parent) :
    QDialog(parent),
	ui(new Ui::PopupWindow)
{
	m_pwi = pwi;

    ui->setupUi(this);

	ui->lblText->setOpenExternalLinks(true);
	setAttribute(Qt::WA_TranslucentBackground);

	ui->lblText->setText(pwi.text);

	if (!pwi.avatar.isNull()) {
		ui->lblAvatar->setScaledContents(true);
		ui->lblAvatar->setPixmap(pwi.avatar);
    }

    ui->lblText->resize(ui->lblText->width(), 60);

	setWindowFlags(Qt::Tool /*| Qt::CustomizeWindowHint*/ | Qt::FramelessWindowHint | Qt::WindowStaysOnTopHint);

    connect(&m_timer, &QTimer::timeout, this, &PopupWindow::onTimer);

	unsigned int nDesktopHeight;
	unsigned int nDesktopWidth;
	unsigned int nScreenWidth;
	unsigned int nScreenHeight;

	QDesktopWidget desktop;
	QRect rcScreen = desktop.screenGeometry(this);
	QRect rcDesktop = desktop.availableGeometry(this);

	nDesktopWidth=rcDesktop.width();
	nDesktopHeight=rcDesktop.height();
	nScreenWidth=rcScreen.width();
	nScreenHeight=rcScreen.height();

	bool bTaskbarOnRight=nDesktopWidth<nScreenWidth && rcDesktop.left()==0;
	bool bTaskbarOnLeft=nDesktopWidth<nScreenWidth && rcDesktop.left()!=0;
	bool bTaskBarOnTop=nDesktopHeight<nScreenHeight && rcDesktop.top()!=0;
	//bool bTaskbarOnBottom=nDesktopHeight<nScreenHeight && rcDesktop.top()==0;

	int nTimeToShow = TIME_TO_SHOW;
	int nTimerDelay;

	m_nIncrement = 2;

	if (bTaskbarOnRight)
	{
		m_nStartPosX=(rcDesktop.right()-m_nLastWindowPosition*width());
		m_nStartPosY=rcDesktop.bottom()-height();
		m_nTaskbarPlacement=TASKBAR_ON_RIGHT;
		nTimerDelay = nTimeToShow/(width()/m_nIncrement);
	}
	else if (bTaskbarOnLeft)
	{
		m_nStartPosX=(rcDesktop.left()-width()+m_nLastWindowPosition*width());
		m_nStartPosY=rcDesktop.bottom()-height();
		m_nTaskbarPlacement=TASKBAR_ON_LEFT;
		nTimerDelay=nTimeToShow/(width()/m_nIncrement);
	}
	else if (bTaskBarOnTop)
	{
		m_nStartPosX=rcDesktop.right()-width();
		m_nStartPosY=(rcDesktop.top()-height()+m_nLastWindowPosition*height());
		m_nTaskbarPlacement=TASKBAR_ON_TOP;
		nTimerDelay=nTimeToShow/(height()/m_nIncrement);
	}
	else //if (bTaskbarOnBottom)
	{
		// Taskbar is on the bottom or Invisible
		m_nStartPosX=rcDesktop.right()-width();
		m_nStartPosY=(rcDesktop.bottom()-m_nLastWindowPosition*height());
		m_nTaskbarPlacement=TASKBAR_ON_BOTTOM;
		nTimerDelay=nTimeToShow/(height()/m_nIncrement);
	}

	m_nCurrentPosX=m_nStartPosX;
	m_nCurrentPosY=m_nStartPosY;

	move(m_nCurrentPosX, m_nCurrentPosY);

	m_nLastWindowPosition++;

	m_bAppearing = true;
	m_timer.setInterval(nTimerDelay);
	m_timer.start();
}

PopupWindow::~PopupWindow()
{
    delete ui;
}

void PopupWindow::changeEvent(QEvent *e)
{
    QDialog::changeEvent(e);
    switch (e->type()) {
    case QEvent::LanguageChange:
        ui->retranslateUi(this);
        break;
    default:
        break;
    }
}

void PopupWindow::onPopupTimeout()
{
	if (isVisible())
		m_timer.start();
}

void PopupWindow::startPopupWaitingTimer()
{
	m_bAppearing = false;
	m_timer.stop();

	int time2live = TIME_TO_LIVE;

    if (this->m_pwi.type == PWPhoneCall)
        time2live = global::getSettingsValue("call_popup_duration", "popup", "5").toInt() * 1000;

    QTimer::singleShot(time2live, this, SLOT(onPopupTimeout()));
}

void PopupWindow::closeAndDestroy()
{
    if (m_PopupWindows.last() == this)
        m_nLastWindowPosition = m_PopupWindows.count() - 1;

	hide();
	m_timer.stop();
	m_PopupWindows.removeOne(this);
	delete this;
	/*close();
	deleteLater();*/
}

void PopupWindow::onTimer()
{
    if (m_bAppearing) // APPEARING
    {
		switch(m_nTaskbarPlacement)
		{
		case TASKBAR_ON_BOTTOM:
			if (m_nCurrentPosY>(m_nStartPosY-height()))
				m_nCurrentPosY-=m_nIncrement;
			else
                startPopupWaitingTimer();
			break;
		case TASKBAR_ON_TOP:
			if ((m_nCurrentPosY-m_nStartPosY)<height())
				m_nCurrentPosY+=m_nIncrement;
			else
                startPopupWaitingTimer();
			break;
		case TASKBAR_ON_LEFT:
			if ((m_nCurrentPosX-m_nStartPosX)<width())
				m_nCurrentPosX+=m_nIncrement;
			else
                startPopupWaitingTimer();
			break;
		case TASKBAR_ON_RIGHT:
			if (m_nCurrentPosX>(m_nStartPosX-width()))
				m_nCurrentPosX-=m_nIncrement;
			else
                startPopupWaitingTimer();
			break;
		}
	}

    else // DISSAPPEARING
    {
		switch(m_nTaskbarPlacement)
		{
		case TASKBAR_ON_BOTTOM:
			if (m_nCurrentPosY<m_nStartPosY) {
				m_nCurrentPosY+=m_nIncrement;
			} else {
                closeAndDestroy();
				return;
			}
			break;
		case TASKBAR_ON_TOP:
			if (m_nCurrentPosY>m_nStartPosY) {
				m_nCurrentPosY-=m_nIncrement;
			} else {
                closeAndDestroy();
				return;
			}
			break;
		case TASKBAR_ON_LEFT:
			if (m_nCurrentPosX>m_nStartPosX) {
				m_nCurrentPosX-=m_nIncrement;
			} else {
                closeAndDestroy();
				return;
			}
			break;
		case TASKBAR_ON_RIGHT:
			if (m_nCurrentPosX<m_nStartPosX) {
				m_nCurrentPosX+=m_nIncrement;
			} else {
                closeAndDestroy();
				return;
			}
			break;
		}
	}

	move(m_nCurrentPosX, m_nCurrentPosY);
}

void PopupWindow::showCallNotification(QString caller)
{
	PWInformation pwi;
	pwi.type = PWPhoneCall;
	pwi.text = tr("You have a call from:<br><b>%1</b>").arg(caller);
    QPixmap avatar;

    if (avatar.isNull())
        avatar = QPixmap(":/images/outcall-logo.png");

	PopupWindow *popup = new PopupWindow(pwi);
	popup->show();
	m_PopupWindows.append(popup);
}

void PopupWindow::showInformationMessage(QString caption, QString message, QPixmap avatar, PWType type)
{
	PWInformation pwi;
	pwi.type = type;
	if (caption!="")
		pwi.text = tr("<b>%1</b><br>%2").arg(caption).arg(message);
	else
		pwi.text = message;

	if (avatar.isNull())
        avatar = QPixmap(":/images/outcall-logo.png");

	pwi.avatar = avatar;

	PopupWindow *popup = new PopupWindow(pwi);
	popup->show();
	m_PopupWindows.append(popup);
}

void PopupWindow::closeAll()
{
	qDeleteAll(m_PopupWindows);
	m_PopupWindows.clear();
	m_nLastWindowPosition = 0;
}

void PopupWindow::mousePressEvent(QMouseEvent *)
{
    m_bAppearing = false;
    m_timer.start();
}

