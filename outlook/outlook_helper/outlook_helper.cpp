// outlook_helper.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "string.h"

//int _tmain(int argc, _TCHAR* argv[])
int main(int argc, char* argv[])
{
	/*const char* dll_path = DLLPATH;
	const char* db_path = DBPATH;
	const char* log_path = LOGPATH;*/
	/*char dll_path[] = DLLPATH;
	char db_path[] = DBPATH;
	char log_path[] = LOGPATH;
	bool bLoadExchangeContacts = true;*/

	/*for (int i = 1; i < argc; i++)
	{
		printf("\nparams: %s\n", argv[i]);
	}*/

	if (argc < 5)
	{
		printf("Missing parameters.");
		return 0;
	}

	if (strcmp(argv[1], "LoadContacts")==0)
	{
		bool bLoadExchangeContacts = false;
		char db_path[256];
		char log_path[256];

		strcpy(db_path, argv[2]);
		strcpy(log_path, argv[3]);


		if (strcmp (argv[4],"true") == 0)
		{
			bLoadExchangeContacts = true;
			//printf("true");
		}

		LoadContacts(db_path, log_path, bLoadExchangeContacts);
	}
	else if (strcmp(argv[1], "AddNewContact")==0)
	{
		char log_path[256];
		char numbers[256];
		char emails[256];
		char name[256];

		strcpy(log_path, argv[2]);
		strcpy(numbers, argv[3]);
		strcpy(emails, argv[4]);
		strcpy(name, argv[5]);

		//void AddNewOutlookContact(const char *log_file, const char *numbers, const char *emails, const char* name, char* result)
		AddNewContact(log_path, numbers, emails, name);
	}
	else if (strcmp(argv[1], "ViewContact")==0)
	{
		char log_path[256];
		char numbers[256];
		char emails[256];
		char name[256];

		strcpy(log_path, argv[2]);
		strcpy(numbers, argv[3]);
		strcpy(emails, argv[4]);
		strcpy(name, argv[5]);

		//void ViewOutlookContact(const char *log_file, const char *numbers, const char *emails, const char* name, char* result)
		ViewContact(log_path, numbers, emails, name);
	}

	return 0;
}

