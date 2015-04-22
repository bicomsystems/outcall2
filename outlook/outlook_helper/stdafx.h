// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include <stdio.h>
#include <windows.h>


void Log(char*);

void LoadContacts(const char* db_path, const char* log_path, bool bLoadExchangeContacts);
void AddNewContact(const char* log_path, const char* numbers, const char* emails, const char* name);
void ViewContact(const char* log_path, const char* numbers, const char* emails, const char* name);