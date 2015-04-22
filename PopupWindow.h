#ifndef POPUPWINDOW_H
#define POPUPWINDOW_H

#include "Global.h"

#include <QDialog>
#include <QTimer>

namespace Ui {
    class PopupWindow;
}

class PopupWindow : public QDialog
{
    Q_OBJECT

public:
    enum PWType
    {
        PWPhoneCall,
        PWInformationMessage
	};

private:

    /**
     * @brief The PWInformation struct / popup window information
     */
    struct PWInformation
    {
		PWType type;
		QString text;
		QPixmap avatar;
		QString extension;
	};

public:
	PopupWindow(const PWInformation& pwi, QWidget *parent = 0);
    ~PopupWindow();

    static void showCallNotification(QString caller);
    static void showInformationMessage(QString caption, QString message, QPixmap avatar=QPixmap(), PWType type = PWInformationMessage);

    static void closeAll();

protected:
    void changeEvent(QEvent *e);

private slots:
    void onTimer();
    void onPopupTimeout();
    virtual void mousePressEvent(QMouseEvent *evet);

private:
    void startPopupWaitingTimer();
    void closeAndDestroy();

private:
    Ui::PopupWindow *ui;

	int m_nStartPosX, m_nStartPosY, m_nTaskbarPlacement;
	int m_nCurrentPosX, m_nCurrentPosY;
	int m_nIncrement; // px
	bool m_bAppearing;

	QTimer m_timer;
	PWInformation m_pwi;

	static QList<PopupWindow*> m_PopupWindows;
	static int m_nLastWindowPosition;
};

#endif // POPUPWINDOW_H
