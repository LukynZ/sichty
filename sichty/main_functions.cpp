using namespace std;

void create_xls(short,int,wstring);
int inputparms(wstring);
wstring inputwstrparms(wstring);

bool entry() {
	wstring answer;
	
	do {
		answer = inputwstrparms(L"Chcete vytvoøit tabulku pro nový mìsíc? (a/n): ");
	} while (answer != L"a" && answer != L"n");
	
	if (answer == L"a") {
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

	create_xls(month, year, name);

}

void continuing() {

}