// Outlook.h : main header file for the Outlook DLL
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// COutlookApp
// See Outlook.cpp for the implementation of this class
//

class COutlookApp : public CWinApp
{
public:
	COutlookApp();

// Overrides
public:
	virtual BOOL InitInstance();

	DECLARE_MESSAGE_MAP()
};


extern "C" { __declspec(dllexport) void LoadContactsIntoDB(const char*, const char*, char*, bool); }
extern "C" { __declspec(dllexport) void StopLoadingContacts(); }
extern "C" { __declspec(dllexport) void AddNewOutlookContact(const char*, const char*, const char*, char*); }
extern "C" { __declspec(dllexport) void ViewOutlookContact(const char*, const char*, const char*, char*); }