// stdafx.cpp : source file that includes just the standard includes
// Outlook.pch will be the pre-compiled header
// stdafx.obj will contain the pre-compiled type information

#include "stdafx.h"

CString gDBPath;
CStringA gResult;
CString gLogFile;

void Log(char *msg) {
	CStringA ansi(msg);
	Log(ansi);
}

#ifdef UNICODE
void Log(CString msg) {
	CStringA ansi(msg);
	Log(ansi);
}
#endif

void Log(CStringA message) {
	FILE *hLogFile = _tfopen(gLogFile, _T("a"));
	if (hLogFile) {
		time_t now;
		struct tm* currentTime;
		time(&now);
		currentTime = localtime(&now);
		fprintf(hLogFile, "%02d.%02d.%d %02d:%02d - %s\r\n", currentTime->tm_mday,
		currentTime->tm_mon + 1, currentTime->tm_year + 1900, currentTime->tm_hour, currentTime->tm_min, message);
		fclose(hLogFile);
	}
}
