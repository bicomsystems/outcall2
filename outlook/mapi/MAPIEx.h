////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: MAPIEx.cpp
// Description: Windows Extended MAPI class 
//
// Copyright (C) 2005-2006, Noel Dillabough
//
// This source code is free to use and modify provided this notice remains intact and that any enhancements
// or bug fixes are posted to the CodeProject page hosting this class for the community to benefit.
//
// Usage: see the Codeproject article at http://www.codeproject.com/internet/CMapiEx.asp
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////

#ifndef __MAPI_EX_H__
#define __MAPI_EX_H__

#ifdef _WIN32_WCE
#include <cemapi.h>
#include <objidl.h>
#else
/*#include <mapix.h>
#include <objbase.h>*/
#undef BOOKMARK
#define BOOKMARK MAPI_BOOKMARK

#include "mapix.h"
//#include "initguid.h"
#include "mapiguid.h"
#include "mapidefs.h"
#include "mapiutil.h"

#define USES_IID_IABContainer
#define USES_IID_IMailUser
#define USES_IID_IAddrBook
#define INITGUID
#include <objbase.h>
#endif

#define RELEASE(s) if(s!=NULL) { s->Release();s=NULL; }

#define MAPI_NO_CACHE ((ULONG) 0x00000200)
#define MDB_ONLINE ((ULONG) 0x00000100)
#define MAPI_CACHE_ONLY ((ULONG) 0x00004000)

// MAPI_BEST_ACCESS /*MAPI_MODIFY | MAPI_NO_CACHE
#define MAPI_FOLDER_FLAGS		MAPI_BEST_ACCESS // 0x1 (when opening contacts' folder)
// MDB_NO_DIALOG | MAPI_BEST_ACCESS
#define MAPI_MSG_STORE_FLAGS	MAPI_BEST_ACCESS

#include <vector>
#include "MAPIContact.h"
//#include "..\StdAfx.h"

class CppSQLite3DB;

class CMAPIEx
{
public:
	CMAPIEx();
	~CMAPIEx();

	enum { PROP_MESSAGE_FLAGS, PROP_ENTRYID, MESSAGE_COLS };

// Attributes
public:
	static int cm_nMAPICode;

protected:
	IMAPISession* m_pSession;
	LPMDB m_pMsgStore;
	LPMAPIFOLDER m_pFolder;
	LPMAPITABLE m_pContents;
	ULONG m_ulMDBFlags;

public:
	void StopLoadingContacts();
	void LoadOutlookContacts(bool bLoadExchangeContacts);
	HRESULT LoadExchangeContacts(CppSQLite3DB *db, CString &error_message);
	//HRESULT LoadExchangePABContacts(CppSQLite3DB *db, CString &error_message);

	IMAPISession* GetSession() { return m_pSession; }
	static LPCTSTR GetValidString(SPropValue& prop);

	static BOOL Init(BOOL bMultiThreadedNotifcations=TRUE);
	static void Term();
	//HRESULT LoadProfileDetails(std::vector<CString> &vProfiles/*, int &nDefault, bool bInitializeMAPI*/);
	
	BOOL Login();
	void Logout();

	/*static CString FillPrefix(CString prefix, CString number);
	static CString FillReplacement(CString prefix, CString replacement, CString number);*/

// Operations
private:
	BOOL OpenMessageStore(LPCTSTR szStore=NULL,ULONG ulFlags=MAPI_FOLDER_FLAGS);

	void LoadContacts(CppSQLite3DB *db);

	LPMDB GetMessageStore() { return m_pMsgStore; }
	LPMAPIFOLDER GetFolder() { return m_pFolder; }

	LPCTSTR GetProfileName(IMAPISession* m_pSession);
	ULONG GetMessageStoreSupport();

	// use bInternal to specify that MAPIEx keeps track of and subsequently RELEASEs the folder
	// remember to eventually RELEASE returned folders if calling with bInternal=FALSE
	LPMAPIFOLDER OpenInbox(BOOL bInternal=TRUE);
	LPMAPIFOLDER OpenFolder(unsigned long ulFolderID,BOOL bInternal);
	LPMAPIFOLDER OpenSpecialFolder(unsigned long ulFolderID,BOOL bInternal);
	LPMAPIFOLDER OpenContacts(BOOL bInternal=TRUE);

	BOOL GetContents(LPMAPIFOLDER pFolder=NULL);
	int GetRowCount();
	BOOL SortContents(ULONG ulSortParam=TABLE_SORT_ASCEND,ULONG ulSortField=PR_MESSAGE_DELIVERY_TIME);
	BOOL GetNextContact(CMAPIContact& contact);

	static void GetNarrowString(SPropValue& prop,CString& strNarrow);	
	
	//HANDLE hStopOutlookThreadEvent;
	bool m_bStop;

	/*std::vector<CString> m_prefixes;
	std::vector<CString> m_replacements;*/
	void ReplaceNumber(CString &number);

	CString GetExchangeValue(LPSPropValue);

	/*static std::vector<CString> Isolate(CString text, bool bComplex);
	static std::vector<CString> Isolate(CString text, std::vector<CString> prefixVector);
	static bool isDifferent(CString replacement, CString prefix);*/
};

extern CMAPIEx gMapiEx;

#ifndef PR_BODY_HTML
#define PR_BODY_HTML PROP_TAG(PT_TSTRING,0x1013)
#endif

#ifndef STORE_HTML_OK
#define	STORE_HTML_OK ((ULONG)0x00010000)
#endif

#ifndef PR_SMTP_ADDRESS
#define PR_SMTP_ADDRESS PROP_TAG(PT_TSTRING,0x39FE)
#endif

#define PR_IPM_APPOINTMENT_ENTRYID (PROP_TAG(PT_BINARY,0x36D0))
#define PR_IPM_CONTACT_ENTRYID (PROP_TAG(PT_BINARY,0x36D1))
#define PR_IPM_JOURNAL_ENTRYID (PROP_TAG(PT_BINARY,0x36D2))
#define PR_IPM_NOTE_ENTRYID (PROP_TAG(PT_BINARY,0x36D3))
#define PR_IPM_TASK_ENTRYID (PROP_TAG(PT_BINARY,0x36D4))
#define PR_IPM_DRAFTS_ENTRYID (PROP_TAG(PT_BINARY,0x36D7))

#endif
