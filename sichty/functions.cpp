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
		array<wstring,12> months_name { L"Leden",L"�nor",L"B�ezen",L"Duben",L"Kv�ten",L"�erven",L"�ervenec",L"Srpen",L"Z���",L"��jen",L"Listopad",L"Prosinec" };

		if (month == 2) {
			add = 0;
			pages = 14;
		}

		wstring file = name + L" - v�kaz pr�ce " + months_name.at(month - 1) + L" " + to_wstring(year) + L".xls";
		wcout << name << endl;

		book = xlCreateBook();

		if (book) {

			unsigned short next{ 24 };

			for (short i {1}; i <= ((2 * pages) + add); i = (i + 2)) {

				if (i == (2 * pages) + add) {
					page = to_wstring(i);
					next = 0;
				}
				else {
					page = to_wstring(i) + L"-" + to_wstring(i + 1);
				}

				sheet = book->addSheet(page.c_str());

				this->layout(next);
				this->set_formating(next);
				this->fill_main_text(next);
				}

			if (book->save(file.c_str())) {
				wcout << L"Byl vytvo�en excel soubor: " + file << endl;
				wcout << L"Najdete ho v adres��i programu. \a" << endl;
			}
			else {
				wcout << L"Nepoda�ilo se vytvo�it soubor." << endl;
			}
		}
	}

	~create_xls() {
		book->release();
	}

	void set_formating(unsigned short next) {

		// rows / cols

		// 21-23 / A-J
		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			this->basefont(format);
			format->setBorder(BORDERSTYLE_MEDIUM);

			for (unsigned short x{ 0 }; x <= next; x += 24) {
				for (unsigned short i{ 21 }; i <= 23; i++) {
					for (unsigned short y{ 0 }; y <= 9; y++) {
						sheet->setCellFormat(i + x, y, format);
					}
				}
			}
		}
		
		// 7-20 / A
		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			this->datafont(format);
			format->setBorder(BORDERSTYLE_MEDIUM);
			format->setRotation(90);

			for (unsigned short x{ 0 }; x <= next; x += 24) {
				for (unsigned short i{ 7 }; i <= 20; i++) {
					sheet->setCellFormat(i + x, 0, format);
				}
			}
		}

		// 7-20 / B-G
		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			this->datafont(format);
			format->setBorder(BORDERSTYLE_THIN);

			for (unsigned short x{ 0 }; x <= next; x += 24) {
				for (unsigned short i{ 7 }; i <= 20; i++) {
					for (unsigned short y{ 1 }; y <= 6; y++) {
						sheet->setCellFormat(i + x, y, format);
					}
				}
			}
		}

		// 7-20 / H-J
		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			this->border(BORDERSTYLE_THIN, BORDERSTYLE_THIN, BORDERSTYLE_MEDIUM, BORDERSTYLE_MEDIUM, format);
			this->datafont(format);

			for (unsigned short x{ 0 }; x <= next; x += 24) {
				for (unsigned short i{ 7 }; i <= 20; i++) {
					for (unsigned short y{ 7 }; y <= 9; y++) {
						sheet->setCellFormat(i + x, y, format);
					}
				}
			}
		}

		// 5-6 / D-I
		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			this->border(BORDERSTYLE_MEDIUM, BORDERSTYLE_MEDIUM, BORDERSTYLE_THIN, BORDERSTYLE_THIN, format);
			this->basefont(format);

			for (unsigned short x{ 0 }; x <= next; x += 24) {
				for (unsigned short i{ 5 }; i <= 6; i++) {
					for (unsigned short y{ 3 }; y <= 8; y++) {
						sheet->setCellFormat(i + x, y, format);
					}
				}
			}
		}

		// 5-6 / A
		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			this->border(BORDERSTYLE_MEDIUM, BORDERSTYLE_MEDIUM, BORDERSTYLE_MEDIUM, BORDERSTYLE_THIN, format);
			this->basefont(format);
			format->setWrap(true);
			this->setformat(6, 0, format, next);
			this->setformat(5, 0, format, next);
		}
			
		// 5-6 / J
		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			this->border(BORDERSTYLE_MEDIUM, BORDERSTYLE_MEDIUM, BORDERSTYLE_THIN, BORDERSTYLE_MEDIUM, format);
			this->basefont(format);
			this->setformat(5, 9, format, next);
			this->setformat(6, 9, format, next);
		}

		// 5 / B-C
		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			this->border(BORDERSTYLE_MEDIUM, BORDERSTYLE_THIN, BORDERSTYLE_THIN, BORDERSTYLE_THIN, format);
			this->basefont(format);
			this->setformat(5, 1, format, next);
			this->setformat(5, 2, format, next);
		}

		// 6 / B-C
		{
			Format* format = book->addFormat();
			this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
			this->border(BORDERSTYLE_THIN, BORDERSTYLE_MEDIUM, BORDERSTYLE_THIN, BORDERSTYLE_THIN, format);
			this->basefont(format);
			this->setformat(6, 1, format, next);
			this->setformat(6, 2, format, next);
		}

		// 1-4 / A-B
		{
			Format* format = book->addFormat();
			this->align(ALIGNH_RIGHT, ALIGNV_CENTER, format);
			this->basefont(format);
			format->setBorder(BORDERSTYLE_MEDIUM);

			for (unsigned short x{ 0 }; x <= next; x += 24) {
				for (unsigned short i{ 1 }; i <= 4; i++) {
					for (unsigned short y{ 0 }; y <= 1; y++) {
						sheet->setCellFormat(i + x, y, format);
					}
				}
			}
		}

		// 1-4 / C-J
		{
			Format* format = book->addFormat();
			this->align(ALIGNH_LEFT, ALIGNV_TOP, format);
			this->basefont(format);
			format->setBorder(BORDERSTYLE_MEDIUM);

			for (unsigned short x{ 0 }; x <= next; x += 24) {
				for (unsigned short i{ 1 }; i <= 4; i++) {
					for (unsigned short y{ 2 }; y <= 9; y++) {
						sheet->setCellFormat(i + x, y, format);
					}
				}
			}
		}

		// 21-23 / A
		{
			Format* format = book->addFormat();
			this->align(ALIGNH_LEFT, ALIGNV_CENTER, format);
			this->basefont(format);
			format->setBorder(BORDERSTYLE_MEDIUM);	

			for (unsigned short x{ 0 }; x <= next; x += 24) {
				for (unsigned short i{ 21 }; i <= 23; i++) {
					sheet->setCellFormat(i + x, 0, format);
				}
			}
		}

		// 21-22 / E-G
		{
			Format* format = book->addFormat();
			this->align(ALIGNH_RIGHT, ALIGNV_CENTER, format);
			this->basefont(format);
			format->setBorder(BORDERSTYLE_MEDIUM);

			for (unsigned short x{ 0 }; x <= next; x += 24) {
				for (unsigned short i{ 21 }; i <= 22; i++) {
					for (unsigned short y{ 4 }; y <= 6; y++) {
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

	void setformat(unsigned short row, unsigned short col, Format * format, unsigned short next) {
		sheet->setCellFormat(row, col, format);
		sheet->setCellFormat(row + next, col, format);
	}

	void basefont(Format * format) {
		Font * font = book->addFont();
		font->setName(L"Calibri");
		font->setSize(11);
		format->setFont(font);
	}

	void datafont(Format * format) {
		Font * font = book->addFont();
		font->setName(L"Calibri");
		font->setSize(9);
		format->setFont(font);
	}

	void layout(unsigned short next) {
		sheet->setDisplayGridlines(true);

		array<double,10> width { 10.71, 6, 5.57, 6.57, 4, 7.71, 4.29, 8.57, 8.57, 28.71 };

		// setting columns width - cols A-J
		for (unsigned short i{ 0 }; i < width.size(); i++) {
			sheet->setCol(i, i, width.at(i));
		}

		// setting row height - rows 1-4 + bottom
		for (unsigned short x{ 0 }; x <= next; x += 24) {
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
		for (unsigned short x{ 0 }; x <= next; x += 24) {
			for (unsigned short i{ 7 }; i <= 23; i++) {
				sheet->setRow(i+x, 18);
			}
		}

		//merging base than + bottom table (+x)
		for (unsigned int x{ 0 }; x <= next; x += 24) {

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

	void fill_main_text(unsigned short next) {

		for (unsigned short x{ 0 }; x <= next; x += 24) {
			sheet->writeStr(1 + x, 0, L"Jm�no a p��jmen�:  ");
			sheet->writeStr(1 + x, 7, L"Pozn�mky:");
			sheet->writeStr(2 + x, 0, L"Datum:  ");
			sheet->writeStr(3 + x, 0, L"Odd�len�:  ");
			sheet->writeStr(4 + x, 0, L"Spolujezdci:  ");
			sheet->writeStr(4 + x, 7, L"Vedouci:");
			sheet->writeStr(5 + x, 0, L"K�d akce\n(�. OP)");
			sheet->writeStr(5 + x, 1, L"�as");
			sheet->writeStr(6 + x, 1, L"Od");
			sheet->writeStr(6 + x, 2, L"Do");
			sheet->writeStr(5 + x, 3, L"hodiny");
			sheet->writeStr(5 + x, 4, L"St�t");
			sheet->writeStr(5 + x, 5, L"SPZ");
			sheet->writeStr(5 + x, 6, L"�/S");
			sheet->writeStr(5 + x, 7, L"Popis pracovn� �innosti");
			sheet->writeStr(21 + x, 0, L"Hodiny celkem:");
			sheet->writeStr(21 + x, 4, L"Kontroloval:  ");
			sheet->writeStr(22 + x, 0, L"P�est�vky:");
			sheet->writeStr(22 + x, 4, L"Datum:  ");
			sheet->writeStr(23 + x, 0, L"�ist� odpracovan� doba:");
		}

	}
};