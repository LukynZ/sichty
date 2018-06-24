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

	short add {1};
	short pages {15};
	wstring page;
	wstring months_name[12] = { L"Leden",L"�nor",L"B�ezen",L"Duben",L"Kv�ten",L"�erven",L"�ervenec",L"Srpen",L"Z���",L"��jen",L"Listopad",L"Prosinec" };
		
	if (month == 2) {
		add = 0;
		pages = 14;
	}

	wstring file = name + L" - v�kaz pr�ce " + months_name[month-1] + L" " + to_wstring(year) + L".xls";
	wcout << name << endl;

	Book *book = xlCreateBook();
	
	if (book) {
	
		for (short i {1}; i <= ((2 * pages) + add); i = (i + 2)) {

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
			wcout << L"Najdete ho v adres��i programu. \a" << endl;
		}
		else {
			wcout << L"Nepoda�ilo se vytvo�it soubor." << endl;
		}
		book->release();
	}
}