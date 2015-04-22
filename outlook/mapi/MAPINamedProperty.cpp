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
#include "MAPINamedProperty.h"
#include <MapiUtil.h>

/////////////////////////////////////////////////////////////
// CMAPINamedProperty

CMAPINamedProperty::CMAPINamedProperty(IMAPIProp* pItem)
{
	m_pItem=pItem;
}

CMAPINamedProperty::~CMAPINamedProperty()
{
}

// gets a custom outlook property (ie EmailAddress1 of a contact)
BOOL CMAPINamedProperty::GetOutlookProperty(ULONG ulData1,ULONG ulProperty,LPSPropValue& pProp)
{
	if(!m_pItem) return FALSE;
	const GUID guidOutlookEmail1={ulData1, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 };

	MAPINAMEID nameID;
	nameID.lpguid=(GUID*)&guidOutlookEmail1;
	nameID.ulKind=MNID_ID;
	nameID.Kind.lID=ulProperty;

	LPMAPINAMEID lpNameID[1]={ &nameID };
	LPSPropTagArray lppPropTags; 

	HRESULT hr=E_INVALIDARG;
	if(m_pItem->GetIDsFromNames(1,lpNameID,0,&lppPropTags)==S_OK) {
		ULONG ulPropCount;
		hr=m_pItem->GetProps(lppPropTags, CMAPIEx::cm_nMAPICode, &ulPropCount, &pProp);
		MAPIFreeBuffer(lppPropTags);
	}
	return (hr==S_OK);
}

const GUID GUIDPublicStrings={0x00020329, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 };

BOOL CMAPINamedProperty::GetNamedProperty(LPCTSTR szFieldName,LPSPropValue &pProp)
{
	MAPINAMEID nameID;
	nameID.lpguid=(GUID*)&GUIDPublicStrings;
	nameID.ulKind=MNID_STRING;
#ifdef UNICODE
	nameID.Kind.lpwstrName=(LPWSTR)szFieldName;
#else
	WCHAR wszFieldName[256];
	MultiByteToWideChar(CP_ACP,0,szFieldName,-1,wszFieldName,255);
	nameID.Kind.lpwstrName=wszFieldName;
#endif

	LPMAPINAMEID lpNameID[1]={ &nameID };
	LPSPropTagArray lppPropTags;

	HRESULT hr=E_INVALIDARG;
	if(m_pItem->GetIDsFromNames(1,lpNameID,0,&lppPropTags)==S_OK) {
		ULONG ulPropCount;
		hr=m_pItem->GetProps(lppPropTags, CMAPIEx::cm_nMAPICode, &ulPropCount, &pProp);
		MAPIFreeBuffer(lppPropTags);
	}
	return (hr==S_OK);
}

BOOL CMAPINamedProperty::GetNamedProperty(LPCTSTR szFieldName,CString& strField)
{
	LPSPropValue pProp;
	if(GetNamedProperty(szFieldName,pProp)) {
		strField=CMAPIEx::GetValidString(*pProp);
		MAPIFreeBuffer(pProp);
		return TRUE;
	}
	return FALSE;
}

BOOL CMAPINamedProperty::SetOutlookProperty(ULONG ulData1,ULONG ulProperty,LPCTSTR szField)
{
	LPSPropValue pProp;
	if(GetOutlookProperty(ulData1,ulProperty,pProp)) {
		pProp->Value.LPSZ=(TCHAR*)szField;
		HRESULT hr=m_pItem->SetProps(1,pProp,NULL);
		MAPIFreeBuffer(pProp);
		return (hr==S_OK);
	}
	return FALSE;
}

BOOL CMAPINamedProperty::SetOutlookProperty(ULONG ulData1,ULONG ulProperty,int nField)
{
	LPSPropValue pProp;
	if(GetOutlookProperty(ulData1,ulProperty,pProp)) {
		pProp->Value.l=nField;
		HRESULT hr=m_pItem->SetProps(1,pProp,NULL);
		MAPIFreeBuffer(pProp);
		return (hr==S_OK);
	}
	return FALSE;
}

// if bCreate is true, the property will be created if necessary otherwise if not present will return FALSE
BOOL CMAPINamedProperty::SetNamedProperty(LPCTSTR szFieldName,LPCTSTR szField,BOOL bCreate)
{
	MAPINAMEID nameID;
	nameID.lpguid=(GUID*)&GUIDPublicStrings;
	nameID.ulKind=MNID_STRING;
#ifdef UNICODE
	int nFieldType=PT_UNICODE;
	nameID.Kind.lpwstrName=(LPWSTR)szFieldName;
#else
	int nFieldType=PT_STRING8;
	WCHAR wszFieldName[256];
	MultiByteToWideChar(CP_ACP,0,szFieldName,-1,wszFieldName,255);
	nameID.Kind.lpwstrName=wszFieldName;
#endif

	LPMAPINAMEID lpNameID[1]={ &nameID };
	LPSPropTagArray lppPropTags; 

	HRESULT hr=E_INVALIDARG;
	if(m_pItem->GetIDsFromNames(1,lpNameID,bCreate ? MAPI_CREATE : 0,&lppPropTags)==S_OK) {
		SPropValue prop;
		prop.ulPropTag=(lppPropTags->aulPropTag[0]|nFieldType);
		prop.Value.LPSZ=(LPTSTR)szField;
		hr=m_pItem->SetProps(1,&prop,NULL);
		MAPIFreeBuffer(lppPropTags);
	}
	return (hr==S_OK);
}
