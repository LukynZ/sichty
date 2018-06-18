using namespace std;

void create_xls(short,int,wstring);
int inputparms(wstring);
wstring inputwstrparms(wstring);

bool entry() {
	wstring answer;
	
	do {
		answer = inputwstrparms(L"Chcete vytvo�it tabulku pro nov� m�s�c? (a/n): ");
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
		month = inputparms(L"Zadejte ��slo m�s�ce: ");
	} while (month < 1 || month > 12);

	do {
		year = inputparms(L"Zadejte rok: ");
	} while (year != 2018);
	
	do {
		name = inputwstrparms(L"Zadejte va�e cel� jm�no: ");
	} while (name == L"");

	create_xls(month, year, name);

}

void continuing() {

}