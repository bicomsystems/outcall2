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
#include "MAPIContact.h"
#include "MAPIEx.h"
#include "MAPINamedProperty.h"

/////////////////////////////////////////////////////////////
// CMAPIContact

const ULONG PhoneNumberIDs[]={
	PR_PRIMARY_TELEPHONE_NUMBER, PR_BUSINESS_TELEPHONE_NUMBER, PR_HOME_TELEPHONE_NUMBER,
	PR_CALLBACK_TELEPHONE_NUMBER, PR_BUSINESS2_TELEPHONE_NUMBER, PR_MOBILE_TELEPHONE_NUMBER,
	PR_RADIO_TELEPHONE_NUMBER, PR_CAR_TELEPHONE_NUMBER, PR_OTHER_TELEPHONE_NUMBER,
	PR_PAGER_TELEPHONE_NUMBER, PR_PRIMARY_FAX_NUMBER, PR_BUSINESS_FAX_NUMBER,
	PR_HOME_FAX_NUMBER, PR_TELEX_NUMBER, PR_ISDN_NUMBER, PR_ASSISTANT_TELEPHONE_NUMBER,
	PR_HOME2_TELEPHONE_NUMBER, PR_TTYTDD_PHONE_NUMBER, PR_COMPANY_MAIN_PHONE_NUMBER, 0
};

const ULONG NameIDs[]={ PR_DISPLAY_NAME, PR_GIVEN_NAME, PR_MIDDLE_NAME, PR_SURNAME, PR_INITIALS, 0 };

CMAPIContact::CMAPIContact()
{
	m_pUser=NULL;
	m_pMAPI=NULL;
	m_entry.cb=0;
	SetEntryID(NULL);
}

CMAPIContact::~CMAPIContact()
{
	Close();
}

void CMAPIContact::SetEntryID(SBinary* pEntry)
{
	if(m_entry.cb) delete [] m_entry.lpb;
	m_entry.lpb=NULL;

	if(pEntry) {
		m_entry.cb=pEntry->cb;
		if(m_entry.cb) {
			m_entry.lpb=new BYTE[m_entry.cb];
			memcpy(m_entry.lpb,pEntry->lpb,m_entry.cb);
		}
	} else {
		m_entry.cb=0;
	}
}

BOOL CMAPIContact::Open(CMAPIEx* pMAPI,SBinary entry)
{
	Close();
	m_pMAPI=pMAPI;
	ULONG ulObjType;
	HRESULT hres = m_pMAPI->GetSession()->OpenEntry(entry.cb,(LPENTRYID)entry.lpb,NULL,MAPI_BEST_ACCESS,&ulObjType,(LPUNKNOWN*)&m_pUser);
	if (hres!=S_OK) {
		CStringA msg;
		msg.Format("Failed to open Contact entry. Error code: 0x%x", hres);
		Log(msg);
		return FALSE;
	}

	SetEntryID(&entry);
	return TRUE;
}

void CMAPIContact::Close()
{
	SetEntryID(NULL);
	RELEASE(m_pUser);
	m_pMAPI=NULL;
}

HRESULT CMAPIContact::GetProperty(ULONG ulProperty,LPSPropValue& pProp)
{
	ULONG ulPropCount;
	ULONG p[2]={ 1,ulProperty };
	return m_pUser->GetProps((LPSPropTagArray)p, CMAPIEx::cm_nMAPICode, &ulPropCount, &pProp);
}

BOOL CMAPIContact::GetPropertyString(CString& strProperty,ULONG ulProperty)
{
	LPSPropValue pProp=NULL;
	if(GetProperty(ulProperty,pProp)==S_OK) {
		strProperty=CMAPIEx::GetValidString(*pProp);
		MAPIFreeBuffer(pProp);
		return TRUE;
	} else {		
		strProperty=_T("");
		if (pProp)
			MAPIFreeBuffer(pProp);
		return FALSE;
	}
}

BOOL CMAPIContact::GetName(CString& strName,ULONG ulNameID)
{
	ULONG i=0;
	while(NameIDs[i]!=ulNameID && NameIDs[i]>0) i++;
	if(!NameIDs[i]) {
		strName=_T("");
		return FALSE;
	} else {
		return GetPropertyString(strName,ulNameID);
	}
}

int CMAPIContact::GetOutlookEmailID(int nIndex)
{
	ULONG ulProperty[]={ OUTLOOK_EMAIL1, OUTLOOK_EMAIL2, OUTLOOK_EMAIL3 };
	if(nIndex<1 || nIndex>3) return 0;
	return ulProperty[nIndex-1];
}

// uses the built in outlook email fields, OUTLOOK_EMAIL1 etc, minus 1 for ADDR_TYPE and +1 for EmailOriginalDisplayName
BOOL CMAPIContact::GetEmail(CString& strEmail,int nIndex)
{
	strEmail = "";

	ULONG nID=GetOutlookEmailID(nIndex);
	if(!nID) return FALSE;

	LPSPropValue pProp;
	CMAPINamedProperty prop(m_pUser);
	if(prop.GetOutlookProperty(OUTLOOK_DATA1,nID-1,pProp)) {
		CString strAddrType=CMAPIEx::GetValidString(*pProp);
		MAPIFreeBuffer(pProp);
		if(prop.GetOutlookProperty(OUTLOOK_DATA1,nID,pProp)) {
			strEmail=CMAPIEx::GetValidString(*pProp);
			MAPIFreeBuffer(pProp);
			if(strAddrType==_T("EX")) {
				// for EX types we use the original display name (seems to contain the appropriate data)
				if(prop.GetOutlookProperty(OUTLOOK_DATA1,nID+1,pProp)) {
					strEmail=CMAPIEx::GetValidString(*pProp);
					MAPIFreeBuffer(pProp);
				}
			}
			return TRUE;
		}
	}
	return FALSE;
}

BOOL CMAPIContact::GetPhoneNumber(CString& strPhoneNumber,ULONG ulPhoneNumberID)
{
	ULONG i=0;
	while(PhoneNumberIDs[i]!=ulPhoneNumberID && PhoneNumberIDs[i]>0) i++;
	if(!PhoneNumberIDs[i]) {
		strPhoneNumber=_T("");
		return FALSE;
	} else {
		return GetPropertyString(strPhoneNumber,ulPhoneNumberID);
	}
}

BOOL CMAPIContact::GetCompany(CString& strCompany)
{
	return GetPropertyString(strCompany,PR_COMPANY_NAME);
}

BOOL CMAPIContact::GetNamedProperty(LPCTSTR szFieldName,CString& strField)
{
	CMAPINamedProperty prop(m_pUser);
	return prop.GetNamedProperty(szFieldName,strField);
}

BOOL CMAPIContact::SetNamedProperty(LPCTSTR szFieldName,LPCTSTR szField,BOOL bCreate)
{
	CMAPINamedProperty prop(m_pUser);
	return prop.SetNamedProperty(szFieldName,szField,bCreate);
}
