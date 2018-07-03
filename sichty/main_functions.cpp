#include "stdafx.h"
#include <iostream>
#include <string>
#include <sstream>
#include "create_xls.h"
#include "main_functions.h"

using namespace std;

bool entry() {
	wstring answer;

	do {
		answer = inputwstrparms(L"Chcete vytvoøit tabulku pro nový mìsíc? (A/n): ");
	} while (answer != L"a" && answer != L"n" && answer != L"");
	
	if (answer == L"a" || answer == L"") {
		return true;
	}
	else {
		return false;
	}
}

void start() {

	short month;
	int year;
	wstring name;
	
	do {
		month = inputparms(L"Zadejte èíslo mìsíce: ");
	} while (month < 1 || month > 12);

	do {
		year = inputparms(L"Zadejte rok: ");
	} while (year != 2018);
	
	do {
		name = inputwstrparms(L"Zadejte vaše celé jméno: ");
	} while (name == L"");

	create_xls timesheet(month, year, name);

}

int inputparms(wstring question) {
	short num;
	wstring helper;

	wcout << question;
	getline(wcin, helper);
	wstringstream myStream(helper);
	if (myStream >> num) {
		return num;
	}
	else {
		return 0;
	}
}

wstring inputwstrparms(wstring question) {
	wstring answer;

	wcout << question;
	getline(wcin, answer);

	return answer;
}

void continuing() {

}