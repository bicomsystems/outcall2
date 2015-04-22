// Outlook.cpp : Defines the initialization routines for the DLL.
//

#include "stdafx.h"
#include "Outlook.h"
#include "MAPIEx.h"
#include <shlwapi.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// If you have Outlook 2000 installed on your machine, then use the following lines
#import "mso.dll"
#import "msoutl9.olb"
 
using namespace Outlook;

//
//TODO: If this DLL is dynamically linked against the MFC DLLs,
//		any functions exported from this DLL which call into
//		MFC must have the AFX_MANAGE_STATE macro added at the
//		very beginning of the function.
//
//		For example:
//
//		extern "C" BOOL PASCAL EXPORT ExportedFunction()
//		{
//			AFX_MANAGE_STATE(AfxGetStaticModuleState());
//			// normal function body here
//		}
//
//		It is very important that this macro appear in each
//		function, prior to any calls into MFC.  This means that
//		it must appear as the first statement within the 
//		function, even before any object variable declarations
//		as their constructors may generate calls into the MFC
//		DLL.
//
//		Please see MFC Technical Notes 33 and 58 for additional
//		details.
//


// COutlookApp

BEGIN_MESSAGE_MAP(COutlookApp, CWinApp)
END_MESSAGE_MAP()


// COutlookApp construction

COutlookApp::COutlookApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}


// The one and only COutlookApp object

COutlookApp theApp;


// COutlookApp initialization

BOOL COutlookApp::InitInstance()
{
	CWinApp::InitInstance();

	return TRUE;
}

CMAPIEx mapi;

void LoadContactsIntoDB(const char */*path*/, const char */*log_file*/, char *result, bool bLoadExchangeContacts) {
	WCHAR path[MAX_PATH];
	//WCHAR log_file_path[MAX_PATH];

	if (SUCCEEDED(SHGetFolderPathW(NULL, CSIDL_PROFILE, NULL, 0, path))==FALSE) {	
		strcpy(result, "Failed to get home path");
		return;
	}

	gDBPath = CString(path) + "\\outcall\\outcall.db";
	gLogFile = CString(path) + "\\outcall\\outcall.log";

	if (PathFileExists((LPCWSTR)gDBPath)==FALSE) {
		strcpy(result, "DB file not found");
		return;
	}

	gResult = "_ok_"; // no errors

	mapi.Logout();

	if (!mapi.Login()) { // initialize MAPI session
		mapi.Logout();
		Log("Failed to initialize MAPI");
		gResult = "Failed to initialize MAPI";
	}
	
	mapi.LoadOutlookContacts(bLoadExchangeContacts);
	
	mapi.Logout(); // close MAPI session

	strncpy(result, gResult.GetBuffer(), 250);
}

void StopLoadingContacts() {
	mapi.StopLoadingContacts();
}

/** This is thread which will display Outlook's New Contact window,
and reload Outlook contacts after the window is closed **/

void AddNewOutlookContact(const char */*log_file*/, const char *numbers, const char *emails, const char* name, char* result)
{
	WCHAR path[MAX_PATH];
	
	if (SUCCEEDED(SHGetFolderPathW(NULL, CSIDL_PROFILE, NULL, 0, path))==FALSE) {	
		strcpy(result, "Failed to get home path");
		return;
	}

	gLogFile = CString(path) + "\\outcall\\outcall.log";

	//gLogFile = log_file;
	gResult = "Open Failed";
	strncpy(result, gResult.GetBuffer(), 250);

	Log("---------- BEGIN OCW ----------");

	_ApplicationPtr pApp;

	HRESULT hres=CoInitialize(NULL);
	if (FAILED(hres))
	{
		CStringA msg;
		msg.Format("Outlook contact window could not be opened. Failed to initialize COM library. Error code = 0x%x", hres);
		Log(msg);
		Log("---------- END OCW ----------");
		//return 0;
		strncpy(result, gResult.GetBuffer(), 250);
		return;
	}

	hres = pApp.CreateInstance("Outlook.Application");
	if (FAILED(hres))
	{
		CoUninitialize();

		CStringA msg;
		msg.Format("Outlook contact window could not be opened. Failed to create Outlook Instance. Error code = 0x%x", hres);
		Log(msg);
		Log("---------- END OCW ----------");
		//return 0;
		strncpy(result, gResult.GetBuffer(), 250);
		return;
	}

	try {
		LPDISPATCH contactPtr;
		_NameSpacePtr pNameSpace = pApp->GetNamespace("MAPI");
		MAPIFolderPtr pFolder = pNameSpace->GetDefaultFolder(olFolderContacts);
		_ContactItemPtr pItem = pFolder->Items->Add();
		if (pItem) {
			//split numbers
			CString str = numbers;

			int nTokenPos = 0;

			CString strToken = str.Tokenize(_T("|"), nTokenPos);

			while (!strToken.IsEmpty())
			{

				int nSubTokenPos = 0;
				CString strType = strToken.Tokenize(_T("="), nSubTokenPos);
				CString strValue = strToken.Tokenize(_T("="), nSubTokenPos);

				_bstr_t bstNumber(strValue);

				if (strType == "Assistant")
					pItem->put_AssistantTelephoneNumber(bstNumber);
				else if (strType == "Business")
					pItem->put_BusinessTelephoneNumber(bstNumber);
				else if (strType == "Business2")
					pItem->put_Business2TelephoneNumber(bstNumber);
				else if (strType == "BusinessFax")
					pItem->put_BusinessFaxNumber(bstNumber);
				else if (strType == "Callback")
					pItem->put_CallbackTelephoneNumber(bstNumber);
				else if (strType == "Car")
					pItem->put_CarTelephoneNumber(bstNumber);
				else if (strType == "CompanyMain")
					pItem->put_CompanyMainTelephoneNumber(bstNumber);
				else if (strType == "Home")
					pItem->put_HomeTelephoneNumber(bstNumber);
				else if (strType == "Home2")
					pItem->put_Home2TelephoneNumber(bstNumber);
				else if (strType == "HomeFax")
					pItem->put_HomeFaxNumber(bstNumber);
				else if (strType == "ISDN")
					pItem->put_ISDNNumber(bstNumber);
				else if (strType == "Mobile")
					pItem->put_MobileTelephoneNumber(bstNumber);				
				else if (strType == "OtherFax")
					pItem->put_OtherFaxNumber(bstNumber);
				else if (strType == "Pager")
					pItem->put_PagerNumber(bstNumber);
				else if (strType == "Primary")
					pItem->put_PrimaryTelephoneNumber(bstNumber);
				else if (strType == "Radio")
					pItem->put_RadioTelephoneNumber(bstNumber);
				else if (strType == "Telex")
					pItem->put_TelexNumber(bstNumber);
				else if (strType == "TTYTDD")
					pItem->put_TTYTDDTelephoneNumber(bstNumber);
				else
					pItem->put_OtherTelephoneNumber(bstNumber);
				
				strToken = str.Tokenize(_T("|"), nTokenPos);
			}

			//split emails
			str = emails;

			nTokenPos = 0;

			strToken = str.Tokenize(_T("|"), nTokenPos);

			int counter = 1;

			while (!strToken.IsEmpty())
			{
				_bstr_t bstEmail(strToken);

				if (counter == 1)
					pItem->put_Email1Address(bstEmail);
				else if (counter == 2)
					pItem->put_Email2Address(bstEmail);
				else if (counter == 3)
					pItem->put_Email3Address(bstEmail);
				else
					break;
				
				strToken = str.Tokenize(_T("|"), nTokenPos);

				counter++;
			}


			_bstr_t bstFullName(name);
			//_bstr_t bstNumber(number);

			pItem->put_FullName(bstFullName);
			//pItem->put_BusinessTelephoneNumber(bstNumber);

			const _variant_t bDisplay= variant_t(true);
			pItem->Display(bDisplay); /*variant_t(true)*/
			//PostMessage(MainFrame->Handle, MSG_RELOAD_OUTLOOK_CONTACTS, 0, 0);
			gResult = "_ok_";
			strncpy(result, gResult.GetBuffer(), 250);
		}
		else
		{
			Log("Failed to create Outlook Contact.");
		}
	}
	catch (...)//Sysutils::Exception &e)
	{
		//DBG_LOG("Exception raised: " + e.Message);
		Log("Exception raised: " /*+ e.Message*/);
	}

	pApp.Release();
	CoUninitialize();

	Log("---------- END OCW ----------");

	//return 0;
	strncpy(result, gResult.GetBuffer(), 250);
	return;
}

void ViewOutlookContact(const char */*log_file*/, const char *numbers, const char *emails, const char* name, char* result)
{
	WCHAR path[MAX_PATH];
	
	if (SUCCEEDED(SHGetFolderPathW(NULL, CSIDL_PROFILE, NULL, 0, path))==FALSE)
	{	
		strcpy(result, "Failed to get home path");
		return;
	}

	gLogFile = CString(path) + "\\outcall\\outcall.log";
	gResult = "Open Failed";
	strncpy(result, gResult.GetBuffer(), 250);

	Log("---------- BEGIN VIEW ----------");

	_ApplicationPtr pApp;

	HRESULT hres = CoInitialize(NULL);
	if (FAILED(hres))
	{
		CStringA msg;
		msg.Format("Outlook contact window could not be opened. Failed to initialize COM library. Error code = 0x%x", hres);
		Log(msg);
		Log("---------- END VIEW ----------");
		strncpy(result, gResult.GetBuffer(), 250);
		return;
	}

	hres = pApp.CreateInstance("Outlook.Application");
	if (FAILED(hres))
	{
		CoUninitialize();

		CStringA msg;
		msg.Format("Outlook contact window could not be opened. Failed to create Outlook Instance. Error code = 0x%x", hres);
		Log(msg);
		Log("---------- END VIEW ----------");
		strncpy(result, gResult.GetBuffer(), 250);
		return;
	}

	try {
		LPDISPATCH contactPtr;

		_NameSpacePtr pNameSpace	= pApp->GetNamespace("MAPI");
		MAPIFolderPtr pFolder		= pNameSpace->GetDefaultFolder(olFolderContacts);
		_ItemsPtr pItems			= pFolder->GetItems();
		_ContactItemPtr pContact	= pItems->GetFirst();
		int nCount					= pItems->GetCount();

		int i = 0;
		for (i; i <  nCount; ++i)
		{
			if (strcmp(name, (char *) pContact->GetFullName()) == 0)
			{
				const _variant_t bDisplay= variant_t(true);
				pContact->Display(bDisplay);
				gResult = "_ok_";
				strncpy(result, gResult.GetBuffer(), 250);
				break;
			}
			pContact = pItems->GetNext();
		}

		if (i == nCount)
		{
			Log("Failed to View Outlook Contact.");
		}
	}
	catch (...)
	{
		Log("Exception raised: ");
	}

	pApp.Release();
	CoUninitialize();

	Log("---------- END VIEW ----------");

	strncpy(result, gResult.GetBuffer(), 250);

	return;
}
