#include "QOutCallLabel.h"

#include <QMouseEvent>
#include <QPainter>

QOutCallLabel::QOutCallLabel(QWidget * parent, Qt::WindowFlags f) :
    QLabel(parent, f)
{
}

void QOutCallLabel::mouseReleaseEvent(QMouseEvent *ev)
{
    if (ev->button() == Qt::LeftButton)
        emit clicked();
    else
        emit rightClicked();
}

void QOutCallLabel::mousePressEvent(QMouseEvent *ev)
{
    if (ev->button() == Qt::LeftButton)
        emit leftMouseDown();
}
