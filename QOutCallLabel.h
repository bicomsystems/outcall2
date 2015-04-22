#ifndef QOUTCALLLABEL_H
#define QOUTCALLLABEL_H

#include <QLabel>

class QOutCallLabel : public QLabel
{
    Q_OBJECT

public:
    QOutCallLabel(QWidget * parent = 0, Qt::WindowFlags f = 0);

protected:
    virtual void mouseReleaseEvent(QMouseEvent *ev);
    virtual void mousePressEvent(QMouseEvent *ev);

signals:
    void clicked();
    void rightClicked();
    void leftMouseDown();
};

#endif // QOUTCALLLABEL_H
