using namespace libxl;

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

void create_xls(short month, int year, wstring name) {

	short add = 1;
	short pages = 15;
	short months[12] = { 31,28,31,30,31,30,31,31,30,31,30,31 };
	wstring months_name[12] = { L"Leden",L"�nor",L"B�ezen",L"Duben",L"Kv�ten",L"�erven",L"�ervenec",L"Srpen",L"Z���",L"��jen",L"Listopad",L"Prosinec" };
	wstring page;
	wstring file = name + L" - v�kaz pr�ce " + months_name[month-1] + L" " + to_wstring(year) + L".xls";
	wcout << name << endl;

	if (year % 4 != 0) {
		months[1] = 28;
	}
	else {
		months[1] = 29;
	}

	if (months[month-1] % 2 == 0) {
		add = 0;
	}

	if (month == 2) {
		pages = 14;
	}

	Book *book = xlCreateBook();
	
	if (book) {
	
		for (short i = 1; i <= ((2 * pages) + add); i = (i + 2)) {

			if (i == (2 * pages) + add) {
				page = to_wstring(i);
			}
			else {
				page = to_wstring(i) + L"-" + to_wstring(i + 1);
			}

			Sheet *sheet = book->addSheet(page.c_str());
		
		}

		if (book->save(file.c_str())) {
			wcout << L"Byl vytvo�en excel soubor: " + file << endl;
			wcout << L"Najdete ho v adres��i programu." << endl;
		}
		else {
			wcout << L"Nepoda�ilo se vytvo�it soubor." << endl;
		}
		book->release();

	}



}