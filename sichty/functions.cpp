using namespace libxl;

class create_xls {

public:

	Book * book;
	Sheet * sheet;
	Format * format;

	create_xls(short month, int year, wstring name) {

		short add {1};
		short pages {15};
		wstring page;
		array<wstring,12> months_name { L"Leden",L"Únor",L"Bøezen",L"Duben",L"Kvìten",L"Èerven",L"Èervenec",L"Srpen",L"Záøí",L"Øíjen",L"Listopad",L"Prosinec" };

		if (month == 2) {
			add = 0;
			pages = 14;
		}

		wstring file = name + L" - výkaz práce " + months_name.at(month - 1) + L" " + to_wstring(year) + L".xls";
		wcout << name << endl;

		book = xlCreateBook();

		if (book) {

			for (short i {1}; i <= ((2 * pages) + add); i = (i + 2)) {

				if (i == (2 * pages) + add) {
					page = to_wstring(i);
				}
				else {
					page = to_wstring(i) + L"-" + to_wstring(i + 1);
				}

				sheet = book->addSheet(page.c_str());

				this->layout();
				this->set_formating();
				this->fill_main_text();
				}

			if (book->save(file.c_str())) {
				wcout << L"Byl vytvoøen excel soubor: " + file << endl;
				wcout << L"Najdete ho v adresáøi programu. \a" << endl;
			}
			else {
				wcout << L"Nepodaøilo se vytvoøit soubor." << endl;
			}
		}
	}

	~create_xls() {
		book->release();
	}

	void set_formating() {
		
		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			format->setBorder(BORDERSTYLE_MEDIUM);

			for (unsigned short x{ 0 }; x <= 24; x += 24) {

				for (unsigned short i{ 21 }; i <= 23; i++) {
					for (unsigned short y{ 0 }; y <= 9; y++) {
						sheet->setCellFormat(i + x, y, format);
					}
				}
			}
		}
		
		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			format->setBorder(BORDERSTYLE_THIN);

			for (unsigned short x{ 0 }; x <= 24; x += 24) {

				for (unsigned short i{ 7 }; i <= 20; i++) {
					for (unsigned short y{ 1 }; y <= 6; y++) {
						sheet->setCellFormat(i + x, y, format);
					}
				}
			}
		}

		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			this->border(BORDERSTYLE_THIN, BORDERSTYLE_THIN, BORDERSTYLE_MEDIUM, BORDERSTYLE_MEDIUM, format);

			for (unsigned short x{ 0 }; x <= 24; x += 24) {

				for (unsigned short i{ 7 }; i <= 20; i++) {
					for (unsigned short y{ 7 }; y <= 9; y++) {
						sheet->setCellFormat(i + x, y, format);
					}
				}
			}
		}

		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			this->border(BORDERSTYLE_MEDIUM, BORDERSTYLE_MEDIUM, BORDERSTYLE_THIN, BORDERSTYLE_THIN, format);

			for (unsigned short x{ 0 }; x <= 24; x += 24) {

				for (unsigned short i{ 5 }; i <= 6; i++) {
					for (unsigned short y{ 3 }; y <= 8; y++) {
						sheet->setCellFormat(i + x, y, format);
					}
				}
			}
		}


		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			this->border(BORDERSTYLE_MEDIUM, BORDERSTYLE_MEDIUM, BORDERSTYLE_MEDIUM, BORDERSTYLE_THIN, format);
			format->setWrap(true);
			this->setformat(6, 0, format);
			this->setformat(5, 0, format);
		}
			
		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			this->border(BORDERSTYLE_MEDIUM, BORDERSTYLE_MEDIUM, BORDERSTYLE_THIN, BORDERSTYLE_MEDIUM, format);
			this->setformat(6, 9, format);
			this->setformat(5, 9, format);
		}

		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			this->border(BORDERSTYLE_MEDIUM, BORDERSTYLE_THIN, BORDERSTYLE_THIN, BORDERSTYLE_THIN, format);
			this->setformat(5, 1, format);
			this->setformat(5, 2, format);
		}

		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			this->border(BORDERSTYLE_THIN, BORDERSTYLE_MEDIUM, BORDERSTYLE_THIN, BORDERSTYLE_THIN, format);
			this->setformat(6, 1, format);
			this->setformat(6, 2, format);
		}

		{
			Format* format = book->addFormat();
			this->align(ALIGNH_RIGHT, ALIGNV_CENTER, format);
			format->setBorder(BORDERSTYLE_MEDIUM);

			for (unsigned short x{ 0 }; x <= 24; x += 24) {
				for (unsigned short i{ 1 }; i <= 4; i++) {
					for (unsigned short y{ 0 }; y <= 1; y++) {
						sheet->setCellFormat(i + x, y, format);
					}
				}
			}
		}

		{
			Format* format = book->addFormat();
			this->align(ALIGNH_LEFT, ALIGNV_TOP, format);
			format->setBorder(BORDERSTYLE_MEDIUM);

			for (unsigned short x{ 0 }; x <= 24; x += 24) {

				for (unsigned short i{ 1 }; i <= 4; i++) {
					for (unsigned short y{ 2 }; y <= 9; y++) {
						sheet->setCellFormat(i + x, y, format);
					}
				}
			}
		}

	}

	void align(AlignH styleH, AlignV styleV, Format * format) {
		format->setAlignH(styleH);
		format->setAlignV(styleV);
	}

	void border(BorderStyle styleT, BorderStyle styleB, BorderStyle styleL, BorderStyle styleR, Format * format) {
		format->setBorderTop(styleT);
		format->setBorderBottom(styleB);
		format->setBorderLeft(styleL);
		format->setBorderRight(styleR);
	}

	void setformat(unsigned short row, unsigned short col, Format * format) {
		sheet->setCellFormat(row, col, format);
		sheet->setCellFormat(row + 24, col, format);
	}

	void layout() {
		sheet->setDisplayGridlines(true);

		array<double,10> width { 10.71, 6, 5.57, 6.57, 4, 7.71, 4.29, 8.57, 8.57, 28.71 };

		// setting columns width - cols A-J
		for (unsigned short i{ 0 }; i < width.size(); i++) {
			sheet->setCol(i, i, width.at(i));
		}

		// setting row height - rows 1-4 + bottom
		for (unsigned short x{ 0 }; x <= 24; x += 24) {
			for (unsigned short i{ 0 }; i <= 4; i++) {
				sheet->setRow(i+x, 15);
			}
		}

		// setting row height 5 + bottom
		sheet->setRow(5, 18);
		sheet->setRow(29, 18);

		// setting row height 6
		sheet->setRow(6, 15);
		sheet->setRow(30, 15);

		// setting row height - rows 7-23 + bottom
		for (unsigned short x{ 0 }; x <= 24; x += 24) {
			for (unsigned short i{ 7 }; i <= 23; i++) {
				sheet->setRow(i+x, 18);
			}
		}

		//merging base than + bottom table (+x)
		for (unsigned int x{ 0 }; x <= 24; x += 24) {

			// merging 1-3 + H-J
			sheet->setMerge(1 + x, 3 + x, 7, 9);

			// merging 1,2,3,4 + A-B | 1,2,3,4 + C-G
			for (unsigned short i{ 1 }; i <= 4; i++) {
				sheet->setMerge(i + x, i + x, 0, 1);
				sheet->setMerge(i + x, i + x, 2, 6);
			}

			// merging 4 + H-J
			sheet->setMerge(4 + x, 4 + x, 7, 9);

			//merging 5 + B-C
			sheet->setMerge(5 + x, 5 + x, 1, 2);

			//merging 5-6 + D,E,F,G
			for (unsigned short i{ 3 }; i <= 6; i++) {
				sheet->setMerge(5 + x, 6 + x, i, i);
			}

			//merging 5-6 + H-J
			sheet->setMerge(5 + x, 6 + x, 7, 9);

			//merging 5-6 + A
			sheet->setMerge(5 + x, 6 + x, 0, 0);

			//merging 7-20 + A
			sheet->setMerge(7 + x, 20 + x, 0, 0);

			//merging 7-20 + H-J
			for (unsigned short i{ 7 }; i <= 20; i++) {
				sheet->setMerge(i + x, i + x, 7, 9);
			}

			//merging 21,22,23 + A-C | 21,22 + F-G | 21,22 + H-J | 23 + E-J
			for (unsigned short i{ 21 }; i <= 23; i++) {

				sheet->setMerge(i + x, i + x, 0, 2);

				if (i != 23) {
					sheet->setMerge(i + x, i + x, 4, 6);
					sheet->setMerge(i + x, i + x, 7, 9);
				}
				else {
					sheet->setMerge(i + x, i + x, 4, 9);
				}
			}
		}
	}

	void fill_main_text() {

		for (unsigned short x{ 0 }; x <= 24; x += 24) {
			sheet->writeStr(1 + x, 0, L"Jméno a pøíjmení:  ");
			sheet->writeStr(1 + x, 7, L"Poznámky:");
			sheet->writeStr(2 + x, 0, L"Datum:  ");
			sheet->writeStr(3 + x, 0, L"Oddìlení:  ");
			sheet->writeStr(4 + x, 0, L"Spolujezdci:  ");
			sheet->writeStr(4 + x, 7, L"Vedouci:");
			sheet->writeStr(5 + x, 0, L"Kód akce\n(è. OP)");
			sheet->writeStr(5 + x, 1, L"Èas");
			sheet->writeStr(6 + x, 1, L"Od");
			sheet->writeStr(6 + x, 2, L"Do");
			sheet->writeStr(5 + x, 3, L"hodiny");
			sheet->writeStr(5 + x, 4, L"Stát");
			sheet->writeStr(5 + x, 5, L"SPZ");
			sheet->writeStr(5 + x, 6, L"Ø/S");
			sheet->writeStr(5 + x, 7, L"Popis pracovní èinnosti");
			sheet->writeStr(21 + x, 0, L"Hodiny celkem:");
			sheet->writeStr(21 + x, 4, L"Kontroloval:");
			sheet->writeStr(22 + x, 0, L"Pøestávky:");
			sheet->writeStr(22 + x, 4, L"Datum: ");
			sheet->writeStr(23 + x, 0, L"Èistá odpracovaná doba:");
		}

	}
};