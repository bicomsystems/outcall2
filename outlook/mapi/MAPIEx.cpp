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

#include "stdafx.h"
#include "MAPIEx.h"

#include "CPPSqlite3U.h"
//#include <edk.h>

using namespace std;

#ifdef UNICODE
int CMAPIEx::cm_nMAPICode=MAPI_UNICODE;
#else
int CMAPIEx::cm_nMAPICode=0;
#endif

CMAPIEx::CMAPIEx()
{	
	m_pSession=NULL;
	m_pMsgStore=NULL;
	m_pFolder=NULL;
	m_pContents=NULL;

	m_ulMDBFlags = MAPI_FOLDER_FLAGS;
}

CMAPIEx::~CMAPIEx()
{
	Logout();
}

BOOL CMAPIEx::Init(BOOL bMultiThreadedNotifcations)
{
	if(MAPIInitialize(NULL)!=S_OK) {
		return FALSE;
	} else {
		return TRUE;
    }
}

void CMAPIEx::Term()
{
	MAPIUninitialize();
#ifdef _WIN32_WCE
	CoUninitialize();
#endif
}

BOOL CMAPIEx::Login()
{
	if (m_pSession) {
		return true; // already initialized MAPI session
	}

	CString profile; //:theApp.GetProfileString("Settings", "MailProfile", "");
	HRESULT hRes;

	if(MAPIInitialize(NULL)!=S_OK) {
		return FALSE;
	}	

	/*if (profile=="") {
		Log("Using system default Mail profile");
	} else {
		CStringA msg = "Using Mail profile: " + profile;
		Log(msg);
	}*/

	//// logon to MAPI for extracting the address book. This can also be used to open the individual mail boxes.
	if (profile!="" && profile!="Default" && profile!="default") {
        TCHAR prof[100];
        _tcscpy(prof, profile);
#ifdef _UNICODE
		hRes = MAPILogonEx(NULL, prof, NULL, MAPI_EXTENDED | MAPI_NEW_SESSION | MAPI_LOGON_UI | MAPI_ALLOW_OTHERS | MAPI_EXPLICIT_PROFILE | MAPI_UNICODE, &m_pSession);
#else
        hRes = MAPILogonEx(NULL, prof, NULL, MAPI_EXTENDED | MAPI_NEW_SESSION | MAPI_LOGON_UI | MAPI_ALLOW_OTHERS | MAPI_EXPLICIT_PROFILE, &m_pSession);
#endif
	} else {
		// Log In to Default profile

		//MAPI_EXTENDED | MAPI_ALLOW_OTHERS | MAPI_NEW_SESSION | MAPI_LOGON_UI | MAPI_EXPLICIT_PROFILE
#ifdef _UNICODE
		hRes = MAPILogonEx (NULL, NULL, NULL, MAPI_EXTENDED | MAPI_NEW_SESSION | MAPI_ALLOW_OTHERS | MAPI_USE_DEFAULT | MAPI_LOGON_UI | MAPI_UNICODE, &m_pSession);
#else
		hRes = MAPILogonEx (NULL, NULL, NULL, MAPI_EXTENDED | MAPI_NEW_SESSION | MAPI_ALLOW_OTHERS | MAPI_USE_DEFAULT | MAPI_LOGON_UI, &m_pSession); // MAPI_EXTENDED | MAPI_NEW_SESSION | MAPI_NO_MAIL
#endif
	}

	if (hRes!=S_OK) {
		m_pSession = NULL;
		return false;
	} else {
		return true; // OpenMessageStore(); - we open message store later on
	}
}

void CMAPIEx::Logout()
{
	RELEASE(m_pContents);
	m_pContents = NULL;
	RELEASE(m_pFolder);
	m_pFolder = NULL;
	RELEASE(m_pMsgStore);
	m_pMsgStore = NULL;
	if (m_pSession) {
		m_pSession->Logoff(NULL, MAPI_LOGOFF_UI, 0);
	}
	RELEASE(m_pSession);
	m_pSession = NULL;

	MAPIUninitialize();
}

// Opens default message store
BOOL CMAPIEx::OpenMessageStore(LPCTSTR szStore,ULONG ulFlags)
{
	if(!m_pSession) return FALSE;

	m_ulMDBFlags=ulFlags;

	LPSRowSet pRows=NULL;
	const int nProperties=3;
	SizedSPropTagArray(nProperties,Columns)={nProperties,{PR_DISPLAY_NAME, PR_ENTRYID, PR_DEFAULT_STORE}};

	BOOL bResult=FALSE;
	IMAPITable*	pMsgStoresTable;
	if(m_pSession->GetMsgStoresTable(0, &pMsgStoresTable)==S_OK) {
		if(pMsgStoresTable->SetColumns((LPSPropTagArray)&Columns, 0)==S_OK) {
			while(TRUE) {
				pRows = NULL;
				if (pMsgStoresTable->QueryRows(1,0,&pRows)!=S_OK) {
					MAPIFreeBuffer(pRows);
					pRows = NULL;
				}
				else if(pRows->cRows!=1) { // we reached the end of the table
					FreeProws(pRows);
					pRows = NULL;					
				}
				else {
					if(!szStore) { 
						if(pRows->aRow[0].lpProps[2].Value.b) {
							CString strStore=GetValidString(pRows->aRow[0].lpProps[0]);
							Log("Default Message Store: " + strStore);
							bResult=TRUE;
						}
					} else {
						CString strStore=GetValidString(pRows->aRow[0].lpProps[0]);
						if(strStore.Find(szStore)!=-1) bResult=TRUE;
					}
					if(!bResult) {
						FreeProws(pRows);
						pRows = NULL;
						continue;
					}
				}
				break;
			}
			if(bResult) {
				RELEASE(m_pMsgStore);
				HRESULT hres = m_pSession->OpenMsgStore(NULL,pRows->aRow[0].lpProps[1].Value.bin.cb,(ENTRYID*)pRows->aRow[0].lpProps[1].Value.bin.lpb,NULL,
					MAPI_MSG_STORE_FLAGS,&m_pMsgStore);
				bResult=(hres==S_OK);
				if (FAILED(hres)) {
					CStringA msg;
					msg.Format("OpenMsgStore failed. Error code: 0x%x", hres);
					Log(msg);
				}
				if (pRows) {
					FreeProws(pRows);
					pRows = NULL;
				}
			}
		}
		RELEASE(pMsgStoresTable);
	}
	return bResult;
}

LPMAPIFOLDER CMAPIEx::OpenFolder(unsigned long ulFolderID,BOOL bInternal)
{
	if(!m_pMsgStore) return NULL;

	LPSPropValue props=NULL;
	ULONG cValues=0;
	DWORD dwObjType;
	ULONG rgTags[]={ 1, ulFolderID };
	LPMAPIFOLDER pFolder;

	if(m_pMsgStore->GetProps((LPSPropTagArray) rgTags, cm_nMAPICode, &cValues, &props)!=S_OK) return NULL;
	HRESULT hRes = m_pMsgStore->OpenEntry(props[0].Value.bin.cb,(LPENTRYID)props[0].Value.bin.lpb, NULL, m_ulMDBFlags, &dwObjType,(LPUNKNOWN*)&pFolder);
	if (FAILED(hRes)) {
		CStringA msg;
		msg.Format("OpenEntry failed. Error code: 0x%x", hRes);
        Log(msg);
	}
	MAPIFreeBuffer(props);

	if(pFolder && bInternal) {
		RELEASE(m_pFolder);
		m_pFolder=pFolder;
	}
	return pFolder;
}

LPMAPIFOLDER CMAPIEx::OpenSpecialFolder(unsigned long ulFolderID,BOOL bInternal)
{
	LPMAPIFOLDER pInbox=OpenInbox(FALSE);
	if(!pInbox || !m_pMsgStore) return FALSE;

	//m_pMsgStore->get

	LPSPropValue props=NULL;
	ULONG cValues=0;
	DWORD dwObjType;
	ULONG rgTags[]={ 1, ulFolderID };
	LPMAPIFOLDER pFolder;

	if(pInbox->GetProps((LPSPropTagArray) rgTags, cm_nMAPICode, &cValues, &props)!=S_OK) return NULL;
	HRESULT hRes = m_pMsgStore->OpenEntry(props[0].Value.bin.cb,(LPENTRYID)props[0].Value.bin.lpb, NULL, m_ulMDBFlags, &dwObjType,(LPUNKNOWN*)&pFolder);
	if (FAILED(hRes)) {
		CStringA msg;
		if (ulFolderID==PR_IPM_CONTACT_ENTRYID)
			msg.Format("OpenEntry failed (Contacts folder). Error code: 0x%x", hRes);
		else
			msg.Format("OpenEntry failed. Error code: 0x%x", hRes);
        Log(msg);
	}
	MAPIFreeBuffer(props);
	RELEASE(pInbox);

	if(pFolder && bInternal) {
		RELEASE(m_pFolder);
		m_pFolder=pFolder;
	}
	return pFolder;
}

LPMAPIFOLDER CMAPIEx::OpenInbox(BOOL bInternal)
{
	if(!m_pMsgStore) {
		Log("msg store not openeed");
		return NULL;
	}

#ifdef _WIN32_WCE
	return OpenFolder(PR_CE_IPM_INBOX_ENTRYID);
#else
	ULONG cbEntryID;
	LPENTRYID pEntryID=NULL;
	DWORD dwObjType;
	LPMAPIFOLDER pFolder;

	if(m_pMsgStore->GetReceiveFolder(NULL,0,&cbEntryID,&pEntryID,NULL)!=S_OK) return NULL;
	HRESULT hRes = m_pMsgStore->OpenEntry(cbEntryID,pEntryID, NULL, m_ulMDBFlags,&dwObjType,(LPUNKNOWN*)&pFolder);
	if (FAILED(hRes)) {
		CStringA msg;
		msg.Format("OpenEntry failed (Inbox). Error code: 0x%x", hRes);
        Log(msg);
	}
	MAPIFreeBuffer(pEntryID);
#endif

	if(pFolder && bInternal) {
		RELEASE(m_pFolder);
		m_pFolder=pFolder;
	}
	return pFolder;
}

LPMAPIFOLDER CMAPIEx::OpenContacts(BOOL bInternal)
{
	return OpenSpecialFolder(PR_IPM_CONTACT_ENTRYID,bInternal);
}

BOOL CMAPIEx::GetContents(LPMAPIFOLDER pFolder)
{
	if(!pFolder) {
		pFolder=m_pFolder;
		if(!pFolder) return FALSE;
	}
	RELEASE(m_pContents);
	if(pFolder->GetContentsTable(CMAPIEx::cm_nMAPICode,&m_pContents)!=S_OK) return FALSE;

	const int nProperties=MESSAGE_COLS;
	SizedSPropTagArray(nProperties,Columns)={nProperties,{PR_MESSAGE_FLAGS, PR_ENTRYID }};
	BOOL retValue = (m_pContents->SetColumns((LPSPropTagArray)&Columns,TBL_BATCH)==S_OK);

	m_pContents->SeekRow(BOOKMARK_BEGINNING, 0, NULL);

	int count = GetRowCount();
	CString tmp;
	tmp.Format(_T("Contents table contains %d items."), count);
	Log(tmp);

	return retValue;
}

int CMAPIEx::GetRowCount()
{
	ULONG ulCount;
	if(!m_pContents || m_pContents->GetRowCount(0,&ulCount)!=S_OK) return -1;
	return ulCount;
}

BOOL CMAPIEx::SortContents(ULONG ulSortParam,ULONG ulSortField)
{
	if(!m_pContents) return FALSE;

	SizedSSortOrderSet(1, SortColums) = {1, 0, 0, {{ulSortField,ulSortParam}}};
	return (m_pContents->SortTable((LPSSortOrderSet)&SortColums,0)==S_OK);
}

BOOL CMAPIEx::GetNextContact(CMAPIContact& contact)
{
	if(!m_pContents) return FALSE;

	LPSRowSet pRows=NULL;
	BOOL bResult=FALSE;
	while(m_pContents->QueryRows(1,NULL,&pRows)==S_OK) {
		if(pRows->cRows) {
			bResult=contact.Open(this,pRows->aRow[0].lpProps[PROP_ENTRYID].Value.bin);
		}
		FreeProws(pRows);
		pRows = NULL;
		break;
	}
	if (pRows)
		MAPIFreeBuffer(pRows);	

	return bResult;
}

// sometimes the string in prop is invalid, causing unexpected crashes
LPCTSTR CMAPIEx::GetValidString(SPropValue& prop)
{
	LPCTSTR s=prop.Value.LPSZ;
	if(s && !::IsBadStringPtr(s,(UINT_PTR)-1)) return s;
	return NULL;
}

// special case of GetValidString to take the narrow string in UNICODE
void CMAPIEx::GetNarrowString(SPropValue& prop,CString& strNarrow)
{
	LPCTSTR s=GetValidString(prop);
	if(!s) strNarrow=_T("");
	else {
#ifdef UNICODE
		// VS2005 can copy directly
		if(_MSC_VER>=1400) {
			strNarrow=(char*)s;
		} else {
			WCHAR wszWide[256];
			MultiByteToWideChar(CP_ACP,0,(LPCSTR)s,-1,wszWide,255);
			strNarrow=wszWide;
		}
#else
		strNarrow=s;
#endif
	}
}


void CMAPIEx::ReplaceNumber(CString &number) {
	//CString prefix, filledPrefix, filledReplacement;

	//number.Replace(_T(" "), _T(""));

	/*for (int i=0; i<m_prefixes.size(); i++) {
		prefix = m_prefixes[i];
		filledPrefix = FillPrefix(prefix, number);
		filledReplacement = FillReplacement(prefix, m_replacements[i], number);
		if (number.Left(filledPrefix.GetLength())==filledPrefix) {			
			//number = m_replacements[i] + number.Mid(prefix.GetLength());			
			number = filledReplacement + number.Mid(filledPrefix.GetLength());
			break;
		}
	}*/

	/*number.Replace(_T("-"), _T(""));
	number.Replace(_T("("), _T(""));
	number.Replace(_T(")"), _T(""));*/
	number.Replace(_T("'"), _T(""));
}

CString CMAPIEx::GetExchangeValue(LPSPropValue lpDN) {
	if (lpDN) {
		#ifdef UNICODE
		return lpDN->Value.lpszW;
		#else
		return lpDN->Value.lpszA;
		#endif
	} else {
		return CString("");
	}
}

HRESULT CMAPIEx::LoadExchangeContacts(CppSQLite3DB *db, CString &error_message) {

	LPMAPISESSION 	pSession    = m_pSession;
	LPADRBOOK 	  	m_pAddrBook = NULL;

	LPMAPICONTAINER pIABRoot = NULL;
	LPMAPITABLE pHTable = NULL;
	//LPMAILUSER pRecipient = NULL;
	LPSRowSet pRows = NULL;

	//ULONG       cbeid    = 0L;
	LPENTRYID   lpeid    = NULL;

	HRESULT     hRes     = S_OK;
	LPSRowSet   pRow = NULL;
	ULONG       ulObjType;
	ULONG       cRows    = 0L;

	//ULONG		ulFlags = 0L;
	ULONG		cbEID = 0L;

	LPMAPITABLE    	 lpContentsTable = NULL;
	LPMAPICONTAINER  lpGAL = NULL;

	LPSPropValue lpDN = NULL;

	SRestriction sres;
	SPropValue spv;
	LPBYTE lpb = NULL;

	ZeroMemory ( (LPVOID) &sres, sizeof (SRestriction ) );

	CString title, name, middleName, surname, suffix, fullName, company, email;
	CString query;	
	CString strTemp;

	error_message = "";

	if (pSession==NULL) {
		error_message = "Failed to load Exhcange contacts. MAPI Session is NULL.";
		Log("Failed to load Exhcange contacts. MAPI Session is NULL.");
		return S_FALSE;
	}

	//there are total of 25 tags, removed 3 that are not used anymore
	SizedSPropTagArray ( 22, sptCols ) = {
	22,
	PR_ENTRYID,
	PR_DISPLAY_NAME,
	/*PR_GIVEN_NAME,
	PR_SURNAME,
	PR_COMPANY_NAME,*/
	PR_ASSISTANT_TELEPHONE_NUMBER,
	PR_BUSINESS_TELEPHONE_NUMBER,
	PR_BUSINESS2_TELEPHONE_NUMBER,
	PR_BUSINESS_FAX_NUMBER,
	PR_CALLBACK_TELEPHONE_NUMBER,
	PR_CAR_TELEPHONE_NUMBER,
	PR_COMPANY_MAIN_PHONE_NUMBER,
	PR_HOME_TELEPHONE_NUMBER,
	PR_HOME2_TELEPHONE_NUMBER,
	PR_HOME_FAX_NUMBER,
	PR_ISDN_NUMBER,
	PR_MOBILE_TELEPHONE_NUMBER,
	PR_OTHER_TELEPHONE_NUMBER,
	PR_PAGER_TELEPHONE_NUMBER,
	PR_PRIMARY_TELEPHONE_NUMBER,
	PR_RADIO_TELEPHONE_NUMBER,
	PR_TELEX_NUMBER,
	PR_TTYTDD_PHONE_NUMBER,
	PR_EMAIL_ADDRESS,
	PR_SMTP_ADDRESS
	};		
	

	hRes = pSession->OpenAddressBook(NULL,NULL,AB_NO_DIALOG,&m_pAddrBook);
	if (FAILED (hRes))
	{
		if (m_pAddrBook) {
			m_pAddrBook->Release();
		}
		CStringA msg;
		msg.Format("Failed to open address book. Error: %X", hRes);
		Log(msg);
		error_message = msg;
		return hRes;
	}

	// - Open the root container of the Address book by calling
	//   IAddrbook::OpenEntry and passing NULL and 0 as the first 2
	//   params.

	if ( FAILED ( hRes = m_pAddrBook -> OpenEntry ( NULL, 0L, &IID_IABContainer, MAPI_FOLDER_FLAGS,	&ulObjType,	(LPUNKNOWN *) &pIABRoot ) ))
	{
		CStringA msg;
		msg.Format("Failed to open the root container of the Address book. Error: %X", hRes);
		Log(msg);
		error_message = msg;
		goto Quit;
	}
	
	if ( ulObjType != MAPI_ABCONT ) {
		error_message = "Failed: ulObjType != MAPI_ABCONT.";
		Log("Failed: ulObjType != MAPI_ABCONT.");
		goto Quit;
	}

	// - Call IMAPIContainer::GetHierarchyTable to open the Hierarchy
	//   table of the root address book container.
	if ( FAILED ( hRes = pIABRoot -> GetHierarchyTable ( NULL, &pHTable ) ) )
	{
		CStringA msg;
		msg.Format("Failed to open the the Hierarchy table of the Address book root container. Error: %X", hRes);
		Log(msg);
		error_message = msg;
		goto Quit;
	}
	
	// - Call HrQueryAllRows using a restriction for the only row
	//   that contains DT_GLOBAL as the PR_DISPLAY_TYPE property.

	SizedSPropTagArray ( 2, ptaga ) = {2,PR_DISPLAY_TYPE,PR_ENTRYID};

	sres.rt                          = RES_PROPERTY;
	sres.res.resProperty.relop       = RELOP_EQ;
	sres.res.resProperty.ulPropTag   = PR_DISPLAY_TYPE;
	sres.res.resProperty.lpProp      = &spv;

	spv.ulPropTag = PR_DISPLAY_TYPE;
	spv.Value.l   = DT_GLOBAL;

	if ( FAILED ( hRes = HrQueryAllRows ( pHTable, (LPSPropTagArray) &ptaga, &sres, NULL, 0L, &pRows ) ) )
	{
		CStringA msg;
		msg.Format("Failed to call HrQueryAllRows. Error: %X", hRes);
		Log(msg);
		error_message = msg;
		MAPIFreeBuffer(pRows);
		pRows = NULL;
		goto Quit;
	}
	else if (pRows->cRows==0) {
		CStringA msg;
		msg.Format("HrQueryAllRows returned 0 rows.");
		Log(msg);
		error_message = msg;
		FreeProws(pRows);
		pRows = NULL;
		goto Quit;
	}
	
	// - Call IAddrbook::OpenEntry using the entry id (PR_ENTRYID)
	// from the row returned from HrQueryAllRows to get back an
	// IMAPIContainer object.

	cbEID = pRows->aRow->lpProps[1].Value.bin.cb;
	lpb = pRows->aRow->lpProps[1].Value.bin.lpb;

	if ( FAILED ( hRes = m_pAddrBook -> OpenEntry ( cbEID, (LPENTRYID)lpb, NULL, MAPI_FOLDER_FLAGS, &ulObjType, (LPUNKNOWN *)&lpGAL ) ) )
	{
		CStringA msg;
		msg.Format("Failed to open Global Address List. Error: %X", hRes);
		Log(msg);
		error_message = msg;
		goto Quit;
	}
	
	if (FAILED(hRes = lpGAL->GetContentsTable(0L, &lpContentsTable))) {
		CStringA msg;
		msg.Format("Get Contents Table failed. Error: %X", hRes);
		Log(msg);
		error_message = msg;
		goto Quit;
	}


	if (FAILED(hRes = lpContentsTable->SetColumns((LPSPropTagArray)&sptCols, /*TBL_ASYNC*/ TBL_BATCH))) {
		CStringA msg;
		msg.Format("Set Columns failed. Error: %X", hRes);
		Log(msg);
		error_message = msg;
		goto Quit;
	}


	if ( FAILED ( hRes = lpContentsTable -> GetRowCount ( 0, &cRows ) ) ) {
		// for some reason, just the first call to this function will fail after MAPI login (initialization), so try again
		// Release and Get contents table once more

		if ( lpContentsTable )
		{
			lpContentsTable -> Release ( );
			lpContentsTable = NULL;
		}
		
		if (FAILED(hRes = lpGAL->GetContentsTable(0L, &lpContentsTable))) {
			CStringA msg;
			msg.Format("Get Contents Table failed. Error: %X", hRes);
			Log(msg);
			error_message = msg;
			goto Quit;
		}

		if (FAILED(hRes = lpContentsTable->SetColumns((LPSPropTagArray)&sptCols, /*TBL_ASYNC*/ TBL_BATCH))) {
			CStringA msg;
			msg.Format("Set Columns failed. Error: %X", hRes);
			Log(msg);
			error_message = msg;
			goto Quit;
		}
		
		if ( FAILED ( hRes = lpContentsTable -> GetRowCount ( 0, &cRows ) ) ) {
			CStringA msg;
			msg.Format("Get Row Count failed. Error: %X", hRes);			
			Log(msg);
			error_message = msg;
			goto Quit;
		}
	}
	

	try {
		ULONG counter = 0;
		CString numbers;

		CStringA sCount;
		sCount.Format("Global Address List contains %d contacts.", cRows);			
		Log(sCount);

		while (counter<cRows) {

			// Query for 10 rows at a time
			pRow = NULL;
			if ( FAILED ( hRes = lpContentsTable -> QueryRows ( 10, 0L, &pRow ))) {
				CStringA msg;
				msg.Format("QueryRows failed. Error: %X", hRes);
				Log(msg);
				error_message = msg;
				MAPIFreeBuffer(pRow);
				pRow = NULL;
				break;
			}

			if (pRow->cRows==0) {
				FreeProws(pRow);
				pRow = NULL;
				break;
			}

			counter += pRow->cRows;

			for(int x=0; (x < pRow->cRows) && !m_bStop; x++)
			{				
				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues,
				PR_DISPLAY_NAME);
				fullName = GetExchangeValue(lpDN);

				numbers = "";

				fullName.Replace(_T("'"), _T("''"));
				
				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_ASSISTANT_TELEPHONE_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Assistant=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_BUSINESS_TELEPHONE_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Business=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_BUSINESS2_TELEPHONE_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Business2=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_BUSINESS_FAX_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Business Fax=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_CALLBACK_TELEPHONE_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Callback=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_CAR_TELEPHONE_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Car=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_COMPANY_MAIN_PHONE_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Company Main=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_HOME_TELEPHONE_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Home=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_HOME2_TELEPHONE_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Home2=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_HOME_FAX_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Home Fax=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_ISDN_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|ISDN=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_MOBILE_TELEPHONE_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Mobile=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_OTHER_TELEPHONE_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Other=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_PRIMARY_FAX_NUMBER); //PR_OTHER_FAX_PHONE_NUMBER
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Primary Fax=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_PAGER_TELEPHONE_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Pager=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_PRIMARY_TELEPHONE_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Primary Telephone=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_RADIO_TELEPHONE_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Radio=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_TELEX_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Telex=" + strTemp;

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_TTYTDD_PHONE_NUMBER);
				strTemp = GetExchangeValue(lpDN);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|TTYTDD=" + strTemp;

				if (numbers!="")
					numbers = numbers.Mid(1);

				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
				pRow->aRow[x].cValues, PR_SMTP_ADDRESS);
				email = GetExchangeValue(lpDN);

				query = "INSERT INTO contacts VALUES ('" + fullName + "', 'numbers', '" + numbers + "')";

				try {
					db->execDML(query.GetBuffer());
				} catch (CppSQLite3Exception& e) {
					Log(e.errorMessage() + CString(" Query: ") + query);
				}

				if (!email.IsEmpty()) {
					email.Replace(_T("'"), _T("''"));
					query = "INSERT INTO contacts VALUES ('" + fullName + "', 'email', '" + email + "')";
					try {
						db->execDML(query.GetBuffer());
					} catch (CppSQLite3Exception& e) {
						Log(e.errorMessage() + CString(" Query: ") + query);
					}
				}
			}
			if (pRow) {
				FreeProws(pRow);
				pRow = NULL;
			}			
		}
	} catch (CppSQLite3Exception& e) {
		Log(CStringA(e.errorMessage()));
	}	

Quit:
	
	if (pRow) {
		FreeProws(pRow);
		pRow = NULL;
	}

	if (pRows) {
		FreeProws(pRows);
		pRows = NULL;
	}
	
	if ( NULL != pHTable)
	{
		pHTable -> Release ( );
		pHTable = NULL;
	}

	if ( NULL != pIABRoot)
	{
		pIABRoot -> Release ( );
		pIABRoot = NULL;
	}

	if ( NULL != lpGAL)
	{
		lpGAL -> Release ( );
		lpGAL = NULL;
	}

	if ( lpContentsTable )
	{
		lpContentsTable -> Release ( );
		lpContentsTable = NULL;
	}

	m_pAddrBook->Release();
	
	return hRes;
}

void CMAPIEx::LoadContacts(CppSQLite3DB *db) {

	Log("Loading Contacts from the current Message Store...");

	if (OpenContacts() && GetContents()) {
		// sort by name (stored in PR_SUBJECT)

		//SortContents(TABLE_SORT_ASCEND,PR_SUBJECT); - caused bug in QueryRows (would return 0 rows even thought in the contents table there were items)

		CMAPIContact contact;		

		CString title, name, middleName, surname, suffix, fullName, company, strTemp, email;
		CString query, numbers;

		try {

			//db->execDML(_T("DELETE FROM contacts"));
			db->execDML(_T("begin transaction"));

			while (GetNextContact(contact) && !m_bStop) {

				/*contact.GetPropertyString(title, PR_TITLE);
				contact.GetPropertyString(name, PR_GIVEN_NAME);
				contact.GetPropertyString(middleName, PR_MIDDLE_NAME);
				contact.GetPropertyString(surname, PR_SURNAME);
				contact.GetPropertyString(company, PR_COMPANY_NAME);*/
				contact.GetName(fullName);

				/*title.Replace(_T("'"), _T("''"));
				name.Replace(_T("'"), _T("''"));
				middleName.Replace(_T("'"), _T("''"));
				surname.Replace(_T("'"), _T("''"));
				suffix.Replace(_T("'"), _T("''"));
				company.Replace(_T("'"), _T("''"));*/
				fullName.Replace(_T("'"), _T("''"));

				numbers = "";
				email = "";

				contact.GetEmail(email);

				contact.GetPhoneNumber(strTemp, PR_ASSISTANT_TELEPHONE_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Assistant=" + strTemp;
				
				contact.GetPhoneNumber(strTemp, PR_BUSINESS_TELEPHONE_NUMBER);				
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Business=" + strTemp;

				contact.GetPhoneNumber(strTemp, PR_BUSINESS2_TELEPHONE_NUMBER);				
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Business2=" + strTemp;				

				contact.GetPhoneNumber(strTemp, PR_BUSINESS_FAX_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Business Fax=" + strTemp;

				contact.GetPhoneNumber(strTemp, PR_CALLBACK_TELEPHONE_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Callback=" + strTemp;
				
				contact.GetPhoneNumber(strTemp, PR_CAR_TELEPHONE_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Car=" + strTemp;
				
				contact.GetPhoneNumber(strTemp, PR_COMPANY_MAIN_PHONE_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Company=" + strTemp;
				
				contact.GetPhoneNumber(strTemp, PR_HOME_TELEPHONE_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Home=" + strTemp;
				
				contact.GetPhoneNumber(strTemp, PR_HOME2_TELEPHONE_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Home2=" + strTemp;
				
				contact.GetPhoneNumber(strTemp, PR_HOME_FAX_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Home Fax=" + strTemp;
				
				contact.GetPhoneNumber(strTemp, PR_ISDN_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|ISDN=" + strTemp;
				
				contact.GetPhoneNumber(strTemp, PR_MOBILE_TELEPHONE_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Mobile=" + strTemp;
				
				contact.GetPhoneNumber(strTemp, PR_OTHER_TELEPHONE_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Other=" + strTemp;
				
				contact.GetPhoneNumber(strTemp, PR_PRIMARY_FAX_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Primary Fax=" + strTemp;
				
				contact.GetPhoneNumber(strTemp, PR_PAGER_TELEPHONE_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Pager=" + strTemp;
				
				contact.GetPhoneNumber(strTemp, PR_PRIMARY_TELEPHONE_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Primary telephone=" + strTemp;
				
				contact.GetPhoneNumber(strTemp, PR_RADIO_TELEPHONE_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Radio=" + strTemp;
				
				contact.GetPhoneNumber(strTemp, PR_TELEX_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|Telex=" + strTemp;
				
				contact.GetPhoneNumber(strTemp, PR_TTYTDD_PHONE_NUMBER);
				ReplaceNumber(strTemp);
				if (strTemp!="")
					numbers += "|TTYTDD=" + strTemp;

				if (numbers!="")
					numbers = numbers.Mid(1);
					
				query = "INSERT INTO contacts VALUES ('" + fullName + "', 'numbers', '" + numbers + "')";
					
				try {
					db->execDML(query.GetBuffer());
				} catch (CppSQLite3Exception& e) {
					Log(e.errorMessage() + CString(" Query: ") + query);
				}

				// add email to database
				if (!email.IsEmpty()) {
					email.Replace(_T("'"), _T("''"));
					query = "INSERT INTO contacts VALUES ('" + fullName + "', 'email', '" + email + "')";

					try {
						db->execDML(query.GetBuffer());
					} catch (CppSQLite3Exception& e) {
						Log(e.errorMessage() + CString(" Query: ") + query);
					}
				}
			}

			db->execDML(_T("end transaction"));

		} catch (CppSQLite3Exception& e) {
			Log(CStringA(e.errorMessage()));
			db->execDML(_T("end transaction"));
		}

		RELEASE(m_pContents);
		RELEASE(m_pFolder);
	} else {
		Log("Failed to open Contacts folder for the current Message Store. Maybe it does not exist?");
		//gResult = "Failed to open Outlook Contacts folder";
	}
}


void CMAPIEx::LoadOutlookContacts(bool bLoadExchangeContacts) {
	m_bStop = false;

	Log("------------------ LOAD OUTLOOK CONTACTS BEGIN ------------------");
	
	if(!m_pSession) return;

	CppSQLite3DB db;
	db.open(CString(gDBPath));

	db.execDML(_T("DELETE FROM contacts"));

	// Iterate through all Message Stores, and load contacts from each one of them (Exchange, Local etc...)

	LPSRowSet pRows=NULL;
	const int nProperties=3;
	SizedSPropTagArray(nProperties,Columns)={nProperties,{PR_DISPLAY_NAME, PR_ENTRYID, PR_DEFAULT_STORE}};

	BOOL bResult=FALSE;
	IMAPITable*	pMsgStoresTable;
	if(m_pSession->GetMsgStoresTable(0, &pMsgStoresTable)==S_OK) {
		if(pMsgStoresTable->SetColumns((LPSPropTagArray)&Columns, 0)==S_OK) {
			while (!m_bStop) {
				pRows = NULL;
				if (pMsgStoresTable->QueryRows(1,0,&pRows)!=S_OK) {
					MAPIFreeBuffer(pRows);
					pRows = NULL;
				}
				else if(pRows->cRows!=1) { // we reached the end of the table
					FreeProws(pRows);
					pRows = NULL;					
				}
				else {

					CString strStore=GetValidString(pRows->aRow[0].lpProps[0]);

					Log("Opening message store: " + strStore);

					RELEASE(m_pMsgStore);
					HRESULT hres = m_pSession->OpenMsgStore(NULL,pRows->aRow[0].lpProps[1].Value.bin.cb,(ENTRYID*)pRows->aRow[0].lpProps[1].Value.bin.lpb,NULL,
						MAPI_MSG_STORE_FLAGS,&m_pMsgStore);
					bResult=(hres==S_OK);
					
					FreeProws(pRows);
					pRows = NULL;
					
					if (FAILED(hres)) {
						CStringA msg;
						msg.Format("Failed to open message store: %s. Error code: 0x%x", strStore, hres);
						Log(msg);
					} else {
						if (!m_bStop)
							LoadContacts(&db);
					}		

					continue;					
				}
				break;
			}			
		}
		RELEASE(pMsgStoresTable);
	}

	if (!m_bStop && bLoadExchangeContacts) {
		Log("Started Loading Exchange GAL...");
		CString error_message;
		HRESULT hRes;
		int cnt = 0;
		do {
			db.execDML(_T("begin transaction"));
			hRes = LoadExchangeContacts(&db, error_message);
			if (hRes==MAPI_E_END_OF_SESSION) {
				Log("MAPI Session ended (MAPI_E_END_OF_SESSION). Retrying GAL contact lookup...");
				db.execDML(_T("rollback transaction"));				
			} else {
				db.execDML(_T("end transaction"));
			}
			cnt++;
		} while (cnt<3 && hRes==MAPI_E_END_OF_SESSION);
		
		if (hRes!=S_OK) {
			CStringA msg;
			msg.Format("Failed to load Exchange Global Address List. Error code: 0x%x", hRes);
			Log(msg);
		}
		Log("Finished Loading Exchange GAL");
	}

	/*if (!m_bStop && bLoadExchangeContacts) {
		Log("------------------ LOAD EXCHANGE CONTACTS BEGIN ------------------");
		CString error_message;
		HRESULT hRes = LoadExchangePABContacts(&db, error_message);
		if (hRes!=S_OK) {
			CStringA msg;
			msg.Format("Failed to load Exchange Private Address List. Error code: 0x%x", hRes);
			Log(msg);
		}
		Log("------------------ LOAD EXCHANGE CONTACTS END ------------------");
	}*/


	Log("------------------ LOAD OUTLOOK CONTACTS END ------------------");
}

void CMAPIEx::StopLoadingContacts() {
	m_bStop=true;
}

// if I try to use MAPI_UNICODE when UNICODE is defined I get the MAPI_E_BAD_CHARWIDTH 
// error so I force narrow strings here
LPCTSTR CMAPIEx::GetProfileName(IMAPISession* pSession)
{
	if(!pSession) return NULL;

	static CString strProfileName;
	LPSRowSet pRows=NULL;
	const int nProperties=2;
	SizedSPropTagArray(nProperties,Columns)={nProperties,{PR_DISPLAY_NAME_A, PR_RESOURCE_TYPE}};

	IMAPITable*	pStatusTable;
	if(pSession->GetStatusTable(0,&pStatusTable)==S_OK) {
		if(pStatusTable->SetColumns((LPSPropTagArray)&Columns, 0)==S_OK) {
			while(TRUE) {
				pRows = NULL;
				if(pStatusTable->QueryRows(1,0,&pRows)!=S_OK) MAPIFreeBuffer(pRows);
				else if(pRows->cRows!=1) FreeProws(pRows);
				else if(pRows->aRow[0].lpProps[1].Value.ul==MAPI_SUBSYSTEM) {
					strProfileName=(LPSTR)GetValidString(pRows->aRow[0].lpProps[0]);
					FreeProws(pRows);
					continue;
				} else {
					FreeProws(pRows);
					continue;
				}
				break;
			}
		}
		RELEASE(pStatusTable);
	}
	return strProfileName;
}



//HRESULT CMAPIEx::LoadExchangePABContacts(CppSQLite3DB *db, CString &error_message) {
//
//	LPMAPISESSION 	pSession    = m_pSession;
//	LPADRBOOK 	  	m_pAddrBook = NULL;
//
//	LPMAPICONTAINER pIABRoot = NULL;
//	LPMAPITABLE pHTable = NULL;
//	//LPMAILUSER pRecipient = NULL;
//	LPSRowSet pRows = NULL;
//
//	//ULONG       cbeid    = 0L;
//	//LPENTRYID   lpeid    = NULL;
//	HRESULT     hRes     = S_OK;
//	LPSRowSet   pRow = NULL;
//	ULONG       ulObjType;
//	ULONG       cRows    = 0L;
//
//	ULONG		ulFlags = 0L;
//	ULONG		cbEID = 0L;
//	LPENTRYID	lpEID = NULL;
//
//	LPMAPITABLE    	 lpContentsTable = NULL;
//	LPMAPICONTAINER  lpPAB = NULL;
//
//	LPSPropValue lpDN = NULL;
//
//	SRestriction sres;
//	SPropValue spv;
//	LPBYTE lpb = NULL;
//
//	ZeroMemory ( (LPVOID) &sres, sizeof (SRestriction ) );
//
//	CString title, name, middleName, surname, suffix, fullName, company, email;
//	CString query;	
//	CString strTemp;
//
//	error_message = "";
//
//	if (pSession==NULL) {
//		error_message = "Failed to load Exhcange contacts. MAPI Session is NULL.";
//		Log("Failed to load Exhcange contacts. MAPI Session is NULL.");
//		return S_FALSE;
//	}
//
//	//there are total of 25 tags, removed 3 that are not used anymore
//	SizedSPropTagArray ( 22, sptCols ) = {
//	22,
//	PR_ENTRYID,
//	PR_DISPLAY_NAME,
//	/*PR_GIVEN_NAME,
//	PR_SURNAME,
//	PR_COMPANY_NAME,*/
//	PR_ASSISTANT_TELEPHONE_NUMBER,
//	PR_BUSINESS_TELEPHONE_NUMBER,
//	PR_BUSINESS2_TELEPHONE_NUMBER,
//	PR_BUSINESS_FAX_NUMBER,
//	PR_CALLBACK_TELEPHONE_NUMBER,
//	PR_CAR_TELEPHONE_NUMBER,
//	PR_COMPANY_MAIN_PHONE_NUMBER,
//	PR_HOME_TELEPHONE_NUMBER,
//	PR_HOME2_TELEPHONE_NUMBER,
//	PR_HOME_FAX_NUMBER,
//	PR_ISDN_NUMBER,
//	PR_MOBILE_TELEPHONE_NUMBER,
//	PR_OTHER_TELEPHONE_NUMBER,
//	PR_PAGER_TELEPHONE_NUMBER,
//	PR_PRIMARY_TELEPHONE_NUMBER,
//	PR_RADIO_TELEPHONE_NUMBER,
//	PR_TELEX_NUMBER,
//	PR_TTYTDD_PHONE_NUMBER,
//	PR_EMAIL_ADDRESS,
//	PR_SMTP_ADDRESS
//	};		
//
//	hRes = pSession->OpenAddressBook(NULL,NULL,AB_NO_DIALOG,&m_pAddrBook);
//	if (FAILED (hRes))
//	{
//		if (m_pAddrBook) {
//			m_pAddrBook->Release();
//		}
//		CStringA msg;
//		msg.Format("Failed to open address book. Error: %X", hRes);
//		Log(msg);
//		error_message = msg;
//		return hRes;
//	}
//
//	/*if ( FAILED ( hRes = HrFindExchangeGlobalAddressList ( m_pAddrBook,
//				&cbeid,
//				&lpeid ) ) )
//	{
//		CStringA msg;
//		msg.Format("Failed to find Exchange Global Address List. Error: %X", hRes);
//		Log(msg);		
//		if (hRes!=0x80004005) // Email account does not use Exchange Server, so don't set error message in that case
// 			error_message = msg;			
//		goto Quit;
//	}*/
//
//	/*if(FAILED(hRes = m_pAddrBook->OpenEntry((ULONG) cbeid,
//					(LPENTRYID) lpeid,
//					NULL,
//					MAPI_BEST_ACCESS | MAPI_NO_CACHE /*MAPI_CACHE_ONLY*//*,
//					&ulObjType,
//					(LPUNKNOWN *)&lpPAB)))
//	{
//		CStringA msg;
//		msg.Format("Failed to open Global Address List. Error: %X", hRes);
//		Log(msg);
//		error_message = msg;
//		goto Quit;
//	}*/
//
//
//	// - Open the root container of the Address book by calling
//	//   IAddrbook::OpenEntry and passing NULL and 0 as the first 2
//	//   params.
//
//	ulFlags = MAPI_BEST_ACCESS | MAPI_NO_CACHE; // MAPI_ACCESS_READ; // MAPI_MODIFY;
//
//	if ( FAILED ( hRes = m_pAddrBook -> OpenEntry ( NULL,	0L,	&IID_IABContainer,	ulFlags,	&ulObjType,	(LPUNKNOWN *) &pIABRoot ) ))
//	{
//		CStringA msg;
//		msg.Format("Failed to open the root container of the Address book. Error: %X", hRes);
//		Log(msg);
//		error_message = msg;
//		goto Quit;
//	}
//	else
//	{
//		ulFlags = 0L;
//	}
//
//	if ( ulObjType != MAPI_ABCONT ) {
//		error_message = "Failed: ulObjType != MAPI_ABCONT.";
//		Log("Failed: ulObjType != MAPI_ABCONT.");
//		goto Quit;
//	}
//
//	// - Call IMAPIContainer::GetHierarchyTable to open the Hierarchy
//	//   table of the root address book container.
//	if ( FAILED ( hRes = pIABRoot -> GetHierarchyTable ( ulFlags, &pHTable ) ) )
//	{
//		CStringA msg;
//		msg.Format("Failed to open the the Hierarchy table of the Address book root container. Error: %X", hRes);
//		Log(msg);
//		error_message = msg;
//		goto Quit;
//	}
//	else
//	{
//		ulFlags = 0l;
//	}
//
//	// - Get Private Address Book Entry ID parameters
//	if ( FAILED ( hRes = m_pAddrBook->GetPAB(&cbEID, &lpEID) ) )
//	{
//		CStringA msg;
//		msg.Format("Couldn't get the Private Address book entry ID. Error: %X", hRes);
//		Log(msg);
//		error_message = msg;
//		goto Quit;
//	}
//
//	//Open Private Address Book using above parameters
//	if ( FAILED( hRes = m_pAddrBook->OpenEntry(cbEID, lpEID, NULL, MAPI_BEST_ACCESS, &ulObjType, (LPUNKNOWN *)&lpPAB) ) )
//	{
//		CStringA msg;
//		msg.Format("Couldn't get the Private Address book. Error: %X", hRes);
//		Log(msg);
//		error_message = msg;
//		goto Quit;
//	}
//
//	if (FAILED(hRes = lpPAB->GetContentsTable(0L, &lpContentsTable))) {
//		CStringA msg;
//		msg.Format("Get Contents Table failed. Error: %X", hRes);
//		Log(msg);
//		error_message = msg;
//		goto Quit;
//	}
//
//
//	if (FAILED(hRes = lpContentsTable->SetColumns((LPSPropTagArray)&sptCols, /*TBL_ASYNC*/ TBL_BATCH))) {
//		CStringA msg;
//		msg.Format("Set Columns failed. Error: %X", hRes);
//		Log(msg);
//		error_message = msg;
//		goto Quit;
//	}
//
//
//	if ( FAILED ( hRes = lpContentsTable -> GetRowCount ( 0, &cRows ) ) ) {
//		// for some reason, just the first call to this function will fail after MAPI login (initialization), so try again
//		// Release and Get contents table once more
//
//		if ( lpContentsTable )
//		{
//			lpContentsTable -> Release ( );
//			lpContentsTable = NULL;
//		}
//		
//		if (FAILED(hRes = lpPAB->GetContentsTable(0L, &lpContentsTable))) {
//			CStringA msg;
//			msg.Format("Get Contents Table failed. Error: %X", hRes);
//			Log(msg);
//			error_message = msg;
//			goto Quit;
//		}
//
//		if (FAILED(hRes = lpContentsTable->SetColumns((LPSPropTagArray)&sptCols, /*TBL_ASYNC*/ TBL_BATCH))) {
//			CStringA msg;
//			msg.Format("Set Columns failed. Error: %X", hRes);
//			Log(msg);
//			error_message = msg;
//			goto Quit;
//		}
//		
//		if ( FAILED ( hRes = lpContentsTable -> GetRowCount ( 0, &cRows ) ) ) {
//			CStringA msg;
//			msg.Format("Get Row Count failed. Error: %X", hRes);			
//			Log(msg);
//			error_message = msg;
//			goto Quit;
//		}
//	}
//
//
//	try {
//		ULONG counter = 0;
//		CString numbers;
//
//		while (counter<cRows) {
//
//			// Query for 25 rows at a time
//
//			if ( FAILED ( hRes = lpContentsTable -> QueryRows ( 25, 0L, &pRow ))) {
//				CStringA msg;
//				msg.Format("QueryRows failed. Error: %X", hRes);
//				Log(msg);
//				error_message = msg;
//				break;
//			}
//
//			if (pRow->cRows==0) {
//				FreeProws(pRow);
//				break;
//			}
//
//			counter += pRow->cRows;
//
//			for(int x=0; (x < pRow->cRows) && !m_bStop; x++)
//			{				
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues,
//				PR_DISPLAY_NAME);
//				fullName = GetExchangeValue(lpDN);
//
//				numbers = "";
//
//				/*lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues,
//				PR_GIVEN_NAME);
//				name = GetExchangeValue(lpDN);			
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues,
//				PR_SURNAME);
//				surname = GetExchangeValue(lpDN);
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues,
//				PR_COMPANY_NAME);
//				company = GetExchangeValue(lpDN);*/
//
//				/*name.Replace(_T("'"), _T("''"));
//				surname.Replace(_T("'"), _T("''"));
//				company.Replace(_T("'"), _T("''"));*/
//				fullName.Replace(_T("'"), _T("''"));
//
//				/*if (name=="" && surname=="") {
//					name = fullName;
//					surname = "";
//				}*/
//
//				//query = _T("INSERT INTO Contacts VALUES ('") + title + "', '" + name + "', '" + middleName + "', '" + surname + "', '" + suffix + "', '" + company + "'";
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_ASSISTANT_TELEPHONE_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Assistant=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_BUSINESS_TELEPHONE_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Business=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_BUSINESS2_TELEPHONE_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Business2=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_BUSINESS_FAX_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Business Fax=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_CALLBACK_TELEPHONE_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Callback=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_CAR_TELEPHONE_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Car=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_COMPANY_MAIN_PHONE_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Company Main=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_HOME_TELEPHONE_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Home=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_HOME2_TELEPHONE_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Home2=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_HOME_FAX_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Home Fax=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_ISDN_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|ISDN=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_MOBILE_TELEPHONE_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Mobile=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_OTHER_TELEPHONE_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Other=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_PRIMARY_FAX_NUMBER); //PR_OTHER_FAX_PHONE_NUMBER
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Primary Fax=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_PAGER_TELEPHONE_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Pager=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_PRIMARY_TELEPHONE_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Primary Telephone=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_RADIO_TELEPHONE_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Radio=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_TELEX_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|Telex=" + strTemp;
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_TTYTDD_PHONE_NUMBER);
//				strTemp = GetExchangeValue(lpDN);
//				ReplaceNumber(strTemp);
//				if (strTemp!="")
//					numbers += "|TTYTDD=" + strTemp;
//
//				if (numbers!="")
//					numbers = numbers.Mid(1);
//
//				lpDN = PpropFindProp(pRow->aRow[x].lpProps,
//				pRow->aRow[x].cValues, PR_SMTP_ADDRESS);
//				email = GetExchangeValue(lpDN);
//
//				query = "INSERT INTO contacts VALUES ('" + fullName + "', 'numbers', '" + numbers + "')";
//
//				try {
//					db->execDML(query.GetBuffer());
//				} catch (CppSQLite3Exception& e) {
//					Log(e.errorMessage() + CString(" Query: ") + query);
//				}
//
//				query = "INSERT INTO contacts VALUES ('" + fullName + "', 'email', '" + email + "')";
//
//				try {
//					db->execDML(query.GetBuffer());
//				} catch (CppSQLite3Exception& e) {
//					Log(e.errorMessage() + CString(" Query: ") + query);
//				}
//			}
//			if (pRow) {
//				FreeProws(pRow);
//				pRow = NULL;
//			}			
//		}
//	} catch (CppSQLite3Exception& e) {
//		Log(CStringA(e.errorMessage()));
//	}	
//
//Quit:	
//	
//	if ( NULL != pHTable)
//	{
//		pHTable -> Release ( );
//		pHTable = NULL;
//	}
//
//	if ( NULL != pIABRoot)
//	{
//		pIABRoot -> Release ( );
//		pIABRoot = NULL;
//	}
//
//	if ( NULL != lpPAB)
//	{
//		lpPAB -> Release ( );
//		lpPAB = NULL;
//	}
//
//	if ( lpContentsTable )
//	{
//		lpContentsTable -> Release ( );
//		lpContentsTable = NULL;
//	}
//
//	m_pAddrBook->Release();
//	
//	return hRes;
//}




//HRESULT CMAPIEx::LoadProfileDetails(vector<CString> &vProfiles/*, int &nDefault, bool bInitializeMAPI*/)
//{
//	HRESULT       hr;
//	
//	vProfiles.clear();
//	vProfiles.push_back("Default");
//
//	//hr = MAPIInitialize(NULL);
//	//if (FAILED(hr)) {
//	//	Log("Failed to initialize MAPI.");
//	//	return hr;
//	//}
//
//	//IMAPISession* pSession;	
// //   hr = MAPILogonEx(NULL, NULL, NULL, MAPI_EXTENDED | MAPI_NEW_SESSION | MAPI_USE_DEFAULT ,&pSession);
//	//if (hr==S_OK) {
//	//	CString ProfileName = GetProfileName(pSession);
//	//	vProfiles.push_back(ProfileName);
//	//	RELEASE(pSession);
//	//} else {
// //       Log("MAPILogonEx Failed.");		
//	//}	
//
//	//MAPIUninitialize();
//
//	return S_OK;
//}


//CString CMAPIEx::FillPrefix(CString prefix, CString number)
//{
//	CString new_prefix = _T("");
//
//	if(number == "")
//	{
//		return new_prefix;
//	}
//
//	for(int i = 0; i < prefix.GetLength(); i++)
//	{
//		if(prefix.GetAt(i) == 'X' || prefix.GetAt(i) == 'x')
//		{
//			new_prefix += number.GetAt(i);
//		}
//		else
//		{
//			new_prefix += prefix.GetAt(i);
//		}
//	}
//	
//	return new_prefix;
//}
//
//std::vector<CString> CMAPIEx::Isolate(CString text, bool bComplex)
//{
//	std::vector<CString> result;
//	CString str = _T("");
//	CString xstr = _T("");
//
//	for(int i = 0; i < text.GetLength(); i++)
//	{		
//		if(text.MakeUpper().GetAt(i)!= 'X')
//		{
//			str += text.MakeUpper().GetAt(i);
//			if(xstr !="")
//			{
//				result.push_back(xstr);
//				xstr = _T("");
//			}
//		}
//		else
//		{
//			xstr+= text.MakeUpper().GetAt(i);
//			if(str !="")
//			{
//				result.push_back(str);
//				str = _T("");
//			}
//		}
//	}
//
//	if(str !="")
//	{
//		result.push_back(str);
//	}
//	
//	if(xstr !="")
//	{
//		result.push_back(xstr);
//	}
//
//	if(bComplex)
//	{
//		std::vector<CString> pre;
//
//		for(int i = 0; i < result.size(); i++)
//		{
//			if(result[i].Find('X') == -1)
//			{
//				pre.push_back(result[i]);
//			}
//			else
//			{
//				for(int j = 0; j < result[i].GetLength(); j++)
//				{
//					pre.push_back(CString("X"));
//				}
//			}
//		}
//
//		return pre;
//	}
//	else
//	{
//		return result;
//	}
//}
//
//std::vector<CString> CMAPIEx::Isolate(CString text, std::vector<CString> prefixVector)
//{
//	std::vector<CString> pre, rep, result;
//	CString str = _T("");
//	CString xstr = _T("");
//
//	for(int i = 0; i < prefixVector.size(); i++)
//	{
//		if(prefixVector[i].Find('X') == -1)
//		{
//			pre.push_back(prefixVector[i]);
//		}
//		else
//		{
//			for(int j = 0; j < prefixVector[i].GetLength(); j++)
//			{
//				pre.push_back(CString("X"));
//			}
//		}
//	}
//
//	for(int i = 0; i < text.GetLength(); i++)
//	{
//		if(text.MakeUpper().GetAt(i) != 'X')
//		{
//			str += text.MakeUpper().GetAt(i);
//		}
//		else
//		{
//			xstr= text.MakeUpper().GetAt(i);
//			if(str != "")
//			{
//				rep.push_back(str);
//				str = _T("");
//			}
//			rep.push_back(xstr);
//		}
//	}		
//
//	int nDiff = pre.size() - rep.size();
//	str = _T("");
//	xstr = _T("");
//	int nFirst = 0, nLast = 0;
//	std::vector<CString> vec;
//
//	if(nDiff > 0)
//	{
//		CString temp = rep[0];
//		nLast = temp.GetLength();
//
//		for(int i = nDiff ; i >= 0; i--)
//		{			
//			if(i!=0)
//			{
//				nFirst = nLast - pre[i].GetLength();
//				vec.push_back(temp.Mid(nFirst,nLast));
//				nLast = nFirst;
//			}
//			else
//			{
//				vec.push_back(temp.Left(nFirst));
//			}
//		}
//
//		for(int i = vec.size()-1; i >= 0; i--)
//		{
//			result.push_back(vec[i]);
//		}
//
//		for(int i = 1; i < rep.size(); i++)
//		{
//			result.push_back(rep[i]);
//		}
//
//		return result;
//	}
//	else if(nDiff < 0)
//	{
//		return rep;
//	}
//	else
//	{
//		return rep;
//	}	
//}
//
//
//CString CMAPIEx::FillReplacement(CString prefix, CString replacement, CString number)
//{
//	if(number == "")
//	{
//		return "";
//	}
//
//	std::vector<CString> prefixVector, replacementVector, filledPrefixVector;
//
//	bool bComplex = isDifferent(replacement, prefix);
//
//	std::vector<CString> isolatedPrefix = Isolate(prefix, bComplex);
//	std::vector<CString> isolatedReplacement;
//
//	if(bComplex)
//	{
//		isolatedReplacement = Isolate(replacement, isolatedPrefix);
//	}
//	else
//	{
//		isolatedReplacement = Isolate(replacement, false);
//	}
//
//	CString filledPrefix = FillPrefix(prefix,number);
//	CString fillReplacement = _T("");
//	
//	int nFirst = 0, nCount = 0;
//
//	for(int i = 0; i < isolatedPrefix.size(); i++)
//	{
//		nCount = isolatedPrefix[i].GetLength();
//		filledPrefixVector.push_back(filledPrefix.Mid(nFirst, nCount));
//		nFirst += nCount;
//	}
//
//	int nPos = 0, nInc = 0;
//
//	for(int i = 0; i < isolatedReplacement.size(); i++)
//	{
//		for(int j = 0 + nInc; j < isolatedPrefix.size(); j++)
//		{
//			if(isolatedPrefix[j] == isolatedReplacement[i])
//			{
//				nPos = j;
//				nInc = j+1;
//				break;
//			}
//			else
//			{					
//				nPos=0;
//			}
//		}
//
//		if(nPos!=0)
//		{
//			fillReplacement +=  filledPrefixVector[nPos];
//		}
//		else
//		{
//			if(isolatedReplacement[i].Find('X')==-1)
//			{
//				fillReplacement += isolatedReplacement[i];
//			}
//		}		
//	}
//
//	return fillReplacement;
//}
//
//bool CMAPIEx::isDifferent(CString replacement, CString prefix)
//{
//	int nStart = 0;
//	int result = 1;
//	int nRepCount = -1;
//	int nPreCount = -1;
//
//	while(result != -1)
//	{
//		result = replacement.MakeUpper().Find(_T("X"),nStart);
//
//		nStart = result + 1;
//		nRepCount++;
//	}
//
//	nStart = 0;
//	result = 1;
//	
//	while(result != -1)
//	{
//		result = prefix.MakeUpper().Find(_T("X"),nStart);
//
//		nStart = result + 1;
//		nPreCount++;
//	}
//
//	if(nPreCount == nRepCount)
//	{
//		return false;
//	}
//	else
//	{
//		return false;
//	}
//}
