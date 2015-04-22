// stdafx.cpp : source file that includes just the standard includes
// outlook_helper.pch will be the pre-compiled header
// stdafx.obj will contain the pre-compiled type information

#include "stdafx.h"

// TODO: reference any additional headers you need in STDAFX.H
// and not in this file

//#include <windows.h>
//#include <shlobj.h>
//#include <WinBase.h>
//#include <string>

typedef void (__cdecl *PLOC)(const char*, const char*, char*, bool);
typedef void (__cdecl *PAOC)(const char*, const char*, const char*, const char*, char*);

BOOL Is64BitProcess(void)
{
#if defined(_WIN64)
	return TRUE;   // 64-bit program
#else
	return FALSE;
#endif
}


void LoadContacts(const char* db_path, const char* log_path, bool bLoadExchangeContacts)
{
	HMODULE hOutlookDLLInstance = Is64BitProcess()?LoadLibraryA("Outlook_x64.dll"):LoadLibraryA("Outlook.dll");
	if (!hOutlookDLLInstance) {
		printf("LoadLibrary failed.");
		return;
	}

	PLOC pLOC;

	pLOC = (PLOC) GetProcAddress(hOutlookDLLInstance, "LoadContactsIntoDB");

	if(NULL != pLOC)
	{
		char buff[256];

		pLOC(db_path, log_path, buff, bLoadExchangeContacts);

		printf("%s", buff);
	}
	else
	{
		//error loading library function
		printf("%s", "Error loading library function.");
	}

	FreeLibrary(hOutlookDLLInstance);

	return;
} 

void AddNewContact(const char* log_path, const char* numbers, const char* emails, const char* name)
{
	HMODULE hOutlookDLLInstance = Is64BitProcess()?LoadLibraryA("Outlook_x64.dll"):LoadLibraryA("Outlook.dll");
	if (!hOutlookDLLInstance) {
		printf("LoadLibrary failed.");
		return;
	}

	PAOC pAOC;

	pAOC = (PAOC) GetProcAddress(hOutlookDLLInstance, "AddNewOutlookContact");

	if(NULL != pAOC)
	{
		char buff[256];

		pAOC(log_path, numbers, emails, name, buff);

		printf("%s", buff);
	}
	else
	{
		//error loading library function
		printf("%s", "Error loading library function.");
	}

	FreeLibrary(hOutlookDLLInstance);

	return;
}

void ViewContact(const char* log_path, const char* numbers, const char* emails, const char* name)
{
	HMODULE hOutlookDLLInstance = Is64BitProcess()?LoadLibraryA("Outlook_x64.dll"):LoadLibraryA("Outlook.dll");
	if (!hOutlookDLLInstance) {
		printf("LoadLibrary failed.");
		return;
	}

	PAOC pAOC;

	pAOC = (PAOC) GetProcAddress(hOutlookDLLInstance, "ViewOutlookContact");

	if(NULL != pAOC)
	{
		char buff[256];

		pAOC(log_path, numbers, emails, name, buff);

		printf("%s", buff);
	}
	else
	{
		//error loading library function
		printf("%s", "Error loading library function.");
	}

	FreeLibrary(hOutlookDLLInstance);

	return;
}
