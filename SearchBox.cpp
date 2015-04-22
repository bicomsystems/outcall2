#include "SearchBox.h"

#include <QTreeWidget>
#include <QHeaderView>
#include <QTimer>
#include <QKeyEvent>

SearchBoxCompleter::SearchBoxCompleter(QLineEdit *parent) :
    QObject(parent),
    m_editor(parent)
{
    m_popup = new QTreeWidget;
    m_popup->setWindowFlags(Qt::Popup);
    m_popup->setFocusPolicy(Qt::NoFocus);
    m_popup->setFocusProxy(parent);
    m_popup->setMouseTracking(true);

    m_popup->setUniformRowHeights(true);
    m_popup->setRootIsDecorated(false);
    m_popup->setEditTriggers(QTreeWidget::NoEditTriggers);
    m_popup->setSelectionBehavior(QTreeWidget::SelectRows);
    m_popup->setFrameStyle(QFrame::Box | QFrame::Plain);
    m_popup->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    m_popup->header()->hide();

    m_popup->installEventFilter(this);

    m_timer = new QTimer(this);
    m_timer->setSingleShot(true);
    m_timer->setInterval(500);
    connect(m_timer, SIGNAL(timeout()), SLOT(autoSuggest()));
    connect(m_editor, SIGNAL(textEdited(QString)), m_timer, SLOT(start()));
    connect(m_popup, SIGNAL(itemClicked(QTreeWidgetItem*,int)), this, SLOT(onSelected()));

}

SearchBoxCompleter::~SearchBoxCompleter()
{
    delete m_popup;
}

bool SearchBoxCompleter::eventFilter(QObject *obj, QEvent *ev)
{
    if (obj != m_popup)
        return false;

    if (ev->type() == QEvent::MouseButtonPress)
    {
        m_popup->hide();
        m_editor->setFocus();
        return true;
    }

    if (ev->type() == QEvent::KeyPress)
    {

        bool consumed = false;
        int key = static_cast<QKeyEvent*>(ev)->key();
        switch (key) {
        case Qt::Key_Enter:
        case Qt::Key_Return:
            onSelected();
            consumed = true;

        case Qt::Key_Escape:
            m_editor->setFocus();
            m_popup->hide();
            consumed = true;

        case Qt::Key_Up:
        case Qt::Key_Down:
        case Qt::Key_Home:
        case Qt::Key_End:
        case Qt::Key_PageUp:
        case Qt::Key_PageDown:
            break;

        default:
            m_editor->setFocus();
            m_editor->event(ev);
            m_popup->hide();
            break;
        }

        return consumed;
    }

    return false;
}

void SearchBoxCompleter::showCompletion(const QStringList &choices)
{
    if (choices.isEmpty() || !m_editor->hasFocus())
        return;

    m_popup->setUpdatesEnabled(false);
    m_popup->clear();
    for (int i = 0; i < choices.count(); ++i)
    {
        QTreeWidgetItem * item;
        item = new QTreeWidgetItem(m_popup);
        item->setText(0, choices[i]);
    }
    m_popup->setCurrentItem(m_popup->topLevelItem(0));
    m_popup->resizeColumnToContents(0);
    m_popup->adjustSize();
    m_popup->setUpdatesEnabled(true);

    int h = m_popup->sizeHintForRow(0) * qMin(7, choices.count()) + 3;
    m_popup->resize(m_popup->width(), h);

    m_popup->move(m_editor->mapToGlobal(QPoint(0, m_editor->height())));
    m_popup->setFocus();
    m_popup->show();
}

void SearchBoxCompleter::autoSuggest()
{
    QString str = m_editor->text();
    QString re  = ".*" + str + ".*";
    QStringList choices = m_data.filter(QRegExp(re, Qt::CaseInsensitive));

    showCompletion(choices);
}

void SearchBoxCompleter::preventSuggest()
{
    m_timer->stop();
}

void SearchBoxCompleter::setFilter(QString filter)
{
    m_filter = filter;
}

void SearchBoxCompleter::setData(QStringList data)
{
    m_data = data;
}

void SearchBoxCompleter::onSelected()
{
    m_timer->stop();
    m_popup->hide();
    m_editor->setFocus();

    QTreeWidgetItem *item = m_popup->currentItem();

    if (item) {
        m_editor->setText(item->text(0));
        QMetaObject::invokeMethod(m_editor, "returnPressed");

        emit selected(item->text(0));
    }
}

SearchBox::SearchBox(QWidget *parent)
    : QLineEdit(parent)
{
    m_completer = new SearchBoxCompleter(this);
    connect(m_completer, SIGNAL(selected(QString)), this, SLOT(onSelected(QString)));

    adjustSize();
    setFocus();
}

void SearchBox::setFilter(QString filter)
{
    m_completer->setFilter(filter);
}

void SearchBox::setData(QStringList data)
{
    m_completer->setData(data);
}

void SearchBox::onSelected(QString text)
{
    emit selected(text);
}
