#ifndef SEARCHBOX_H
#define SEARCHBOX_H

#include <QObject>
#include <QLineEdit>

class QTimer;
class QTreeWidget;

class SearchBoxCompleter : public QObject
{
    Q_OBJECT

public:
    SearchBoxCompleter(QLineEdit *parent = 0);
    ~SearchBoxCompleter();

    bool eventFilter(QObject *obj, QEvent *ev);
    void showCompletion(const QStringList &choices);
    void setFilter(QString filter);
    void setData(QStringList data);

protected slots:
    void preventSuggest();
    void autoSuggest();
    void onSelected();

signals:
    void selected(QString text);

private:
    QLineEdit   *m_editor;
    QTreeWidget *m_popup;
    QTimer      *m_timer;
    QStringList m_data;
    QString     m_filter;
};

class SearchBox : public QLineEdit
{
    Q_OBJECT

public:
    SearchBox(QWidget *parent = 0);

    void setFilter(QString filter);
    void setData(QStringList data);

protected slots:
    void onSelected(QString text);

signals:
    void selected(QString text);

private:
    SearchBoxCompleter *m_completer;
};

#endif // SEARCHBOX_H
