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

#ifndef __MAPICONTACT_H__
#define __MAPICONTACT_H__

#undef BOOKMARK
#define BOOKMARK MAPI_BOOKMARK
#include "initguid.h"

#include "mapiguid.h"
#include "mapidefs.h"
#include "mapiutil.h"

#define USES_IID_IABContainer
#define USES_IID_IMailUser
#define USES_IID_IAddrBook
#define INITGUID

//#include <MapiUtil.h>

class CMAPIEx;
class CMAPIContact;

/////////////////////////////////////////////////////////////
// CMAPIContact

class CMAPIContact
{
public:
	CMAPIContact();
	~CMAPIContact();

	enum { OUTLOOK_DATA1=0x00062004, OUTLOOK_EMAIL1=0x8083, OUTLOOK_EMAIL2=0x8093, OUTLOOK_EMAIL3=0x80A3,
		OUTLOOK_FILE_AS=0x8005, OUTLOOK_POSTAL_ADDRESS=0x8022, OUTLOOK_DISPLAY_ADDRESS_HOME=0x801A };

// Attributes
protected:
	CMAPIEx* m_pMAPI;
	LPMAILUSER m_pUser;
	SBinary m_entry;

// Operations
public:
	void SetEntryID(SBinary* pEntry=NULL);
	SBinary* GetEntryID() { return &m_entry; }

	BOOL Open(CMAPIEx* pMAPI,SBinary entry);
	void Close();
	BOOL Save(BOOL bClose=TRUE);

	BOOL GetPropertyString(CString& strProperty,ULONG ulProperty);
	BOOL GetName(CString& strName,ULONG ulNameID=PR_DISPLAY_NAME);
	BOOL GetCompany(CString& strCompany);
	BOOL GetEmail(CString& strEmail,int nIndex=1); // 1, 2 or 3 for outlook email addresses
	BOOL GetPhoneNumber(CString& strPhoneNumber,ULONG ulPhoneNumberID);
	BOOL GetNamedProperty(LPCTSTR szFieldName,CString& strField);

    BOOL SetNamedProperty(LPCTSTR szFieldName,LPCTSTR szField,BOOL bCreate=TRUE);

protected:
	int GetOutlookEmailID(int nIndex);
	HRESULT GetProperty(ULONG ulProperty,LPSPropValue &pProp);
};

#endif
