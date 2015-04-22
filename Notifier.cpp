#include "Notifier.h"

Notifier *g_Notifier = nullptr;

Notifier::Notifier()
{
    g_Notifier = this;
}

Notifier::~Notifier()
{
}

void Notifier::emitSettingsChanged()
{
    emit settingsChanged();
}
