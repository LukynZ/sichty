#include "stdafx.h"
#include <iostream>
#include <string>
#include <array>
#include "create_xls.h"
#include "main_functions.h"


using namespace libxl;

create_xls::create_xls(short month, int year, std::wstring name) {

	short add {1};
	short pages {15};
	std::wstring page;
	std::wstring date_m_y;
	std::wstring answer;
	
	std::array<std::wstring,12> months_name { L"Leden",L"�nor",L"B�ezen",L"Duben",L"Kv�ten",L"�erven",L"�ervenec",L"Srpen",L"Z���",L"��jen",L"Listopad",L"Prosinec" };

	if (month == 2) {
		add = 0;
		pages = 14;
	}

	std::wstring file { name + L" - v�kaz pr�ce " + months_name.at(month - 1) + L" " + std::to_wstring(year) + L".xls" };

	book = xlCreateBook();

	if (book) {

		unsigned short next{ 24 };
		bool fill_next = true;

		for (unsigned short i{ 1 }; i <= ((2 * pages) + add); i = (i + 2)) {

			if (i == (2 * pages) + add) {
				page = std::to_wstring(i);
				next = 0;
			}
			else {
				page = std::to_wstring(i) + L"-" + std::to_wstring(i + 1);
			}

			date_m_y = std::to_wstring(month) + L"." + std::to_wstring(year);
			sheet = book->addSheet(page.c_str());
			this->layout(next);
			this->set_formating(next);
			this->fill_main_text(next, name, i, month, year);

			for (unsigned short x{ i }, y{ 0 }; x <= (i + 1) && x <= ((2 * pages) + add) && fill_next == true; x++, y += 24) {
				do {
					answer = inputwstrparms(L"Chcete p�idat z�znam pro: " + std::to_wstring(x) + L"." + date_m_y + L"? (a/n/nv) *nv = ne vse: ");
				} while (answer != L"a" && answer != L"n" && answer != L"nv");

				if (answer == L"a") {
					this->fill_data(y);
				}
				else if (answer == L"nv") {
					fill_next = false;
				}
				
			}
		}

		if (book->save(file.c_str())) {
			std::wcout << L"Byl vytvo�en excel soubor: " + file << std::endl;
			std::wcout << L"Najdete ho v adres��i programu. \a" << std::endl;
		}
		else {
			std::wcout << L"Nepoda�ilo se vytvo�it soubor." << std::endl;
		}
	}
}

create_xls::~create_xls() {
	book->release();
}

void create_xls::set_formating(unsigned short next) {

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

	// 7-20 / B-D
	{
		Format* format = book->addFormat();
		this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
		this->datafont(format);
		format->setBorder(BORDERSTYLE_THIN);
		format->setNumFormat(NUMFORMAT_CUSTOM_HMM);

		for (unsigned short x{ 0 }; x <= next; x += 24) {
			for (unsigned short i{ 7 }; i <= 20; i++) {
				for (unsigned short y{ 1 }; y <= 3; y++) {
					sheet->setCellFormat(i + x, y, format);
				}
			}
		}
	}

	// 7-20 / E-G
	{
		Format* format = book->addFormat();
		this->align(ALIGNH_CENTER, ALIGNV_CENTER, format);
		this->datafont(format);
		format->setBorder(BORDERSTYLE_THIN);

		for (unsigned short x{ 0 }; x <= next; x += 24) {
			for (unsigned short i{ 7 }; i <= 20; i++) {
				for (unsigned short y{ 4 }; y <= 6; y++) {
					sheet->setCellFormat(i + x, y, format);
				}
			}
		}
	}

	// 7-20 / H-J
	{
		Format* format = book->addFormat();
		this->align(ALIGNH_LEFT, ALIGNV_CENTER, format);
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

	// 21-23 / D
	{
		Format* format = book->addFormat();
		this->align(ALIGNH_LEFT, ALIGNV_CENTER, format);
		this->basefont(format);
		format->setBorder(BORDERSTYLE_MEDIUM);
		format->setNumFormat(NUMFORMAT_CUSTOM_HMM);

		for (unsigned short x{ 0 }; x <= next; x += 24) {
			for (unsigned short i{ 21 }; i <= 23; i++) {
				sheet->setCellFormat(i + x, 3, format);
			}
		}
	}

	// 2 / C
	{
		Format* format = book->addFormat();
		this->align(ALIGNH_LEFT, ALIGNV_CENTER, format);
		this->boldfont(format);
		format->setBorder(BORDERSTYLE_MEDIUM);
		sheet->setCellFormat(2, 2, format);
		sheet->setCellFormat(26, 2, format);
	}

}

void create_xls::align(AlignH styleH, AlignV styleV, Format * format) {
	format->setAlignH(styleH);
	format->setAlignV(styleV);
}

void create_xls::border(BorderStyle styleT, BorderStyle styleB, BorderStyle styleL, BorderStyle styleR, Format * format) {
	format->setBorderTop(styleT);
	format->setBorderBottom(styleB);
	format->setBorderLeft(styleL);
	format->setBorderRight(styleR);
}

void create_xls::setformat(unsigned short row, unsigned short col, Format * format, unsigned short next) {
	sheet->setCellFormat(row, col, format);
	sheet->setCellFormat(row + next, col, format);
}

void create_xls::basefont(Format * format) {
	Font * font = book->addFont();
	font->setName(L"Calibri");
	font->setSize(11);
	format->setFont(font);
}

void create_xls::datafont(Format * format) {
	Font * font = book->addFont();
	font->setName(L"Calibri");
	font->setSize(9);
	format->setFont(font);
}

void create_xls::boldfont(Format * format) {
	Font * font = book->addFont();
	font->setBold(true);
	format->setFont(font);
}

void create_xls::layout(unsigned short next) {
	sheet->setDisplayGridlines(true);

	std::array<double,10> width { 10.71, 6, 5.57, 6.57, 4, 7.71, 4.29, 8.57, 8.57, 28.71 };

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

void create_xls::fill_main_text(unsigned short next, std::wstring name, unsigned short day, unsigned short month, unsigned int year) {

	for (unsigned short x{ 0 }; x <= next; x += 24, day++) {
		std::wstring sheetdate{ L"  " + std::to_wstring(day) + L"." + std::to_wstring(month) + L"." + std::to_wstring(year) };

		sheet->writeStr(1 + x, 0, L"  Jm�no a p��jmen�:  ");
		sheet->writeStr(1 + x, 2, (L"  " + name).c_str());
		sheet->writeStr(1 + x, 7, L"Pozn�mky:");
		sheet->writeStr(2 + x, 0, L"Datum:  ");
		sheet->writeStr(2 + x, 2, sheetdate.c_str());
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

void create_xls::fill_data(unsigned short next) {

	unsigned int num_input_from;
	unsigned int num_input_to;
	unsigned int time;
	unsigned short start;
	std::wstring input;
	std::wstring spz;
	std::wstring country{ L"D" };
	std::wstring activity;
	bool target;

	start = 7 + next;

	std::wcout << L"Zad�vejte �asov� �daje bez dvojte�ky!" << std::endl;

	for (unsigned short i{ start }; i <= (20 + next); i++) {
		target = false;
		for (unsigned short x{ 1 }; x <= 4; x++) {

			switch (x) {

			case 1:
				do {
					num_input_from = inputparms(L"Zadejte �as OD: ");
				} while (num_input_from < 0 || num_input_from > 2400);
				input = this->convert_time(num_input_from);
				sheet->writeStr(i, 1, input.c_str());
				break;

			case 2:
				do {
					num_input_to = inputparms(L"Zadejte �as DO: ");
				} while (num_input_to <= num_input_from || num_input_to > 2400);

				input = this->convert_time(num_input_to);
				sheet->writeStr(i, 2, input.c_str());

				time = num_input_to - num_input_from;
				input = this->convert_time(time);
				sheet->writeStr(i, 3, input.c_str());

				break;

			case 3:
				do {
					input = inputwstrparms(L"Zem� D/CZ/PL? (d/c/p): ");
				}  while (input != L"d" && input != L"c" && input != L"p");
				
				if (input == L"d") {
					input = L"D";
				}
				else if (input == L"c") {
					input = L"CZ";
				}
				else {
					input = L"PL";
				}
				
				sheet->writeStr(i, 4, input.c_str());
				break;

			case 4:
				do {
					input = inputwstrparms(L"�innost? (c)esta, (s)wap, (p)��prava, (z)a�i�t�n�, s(u)rwey, (t)ickets, (d)okumentace, (a)dministrativa: ");
				} while (input != L"c" && input != L"s" && input != L"p" && input != L"z" && input != L"u" && input != L"t" && input != L"d" && input != L"a");

				if (input == L"c") {
					activity = L"-> ";
					if (spz != L"") {
						do {
							input = inputwstrparms(L"SPZ vozidla je " + spz + L"? (a/n) ");
						} while (input != L"a" && input != L"n");

						if (input == L"a") {
							sheet->writeStr(i, 5, spz.c_str());
						}
						else {
							do {
								spz = inputwstrparms(L"Zadejte SPZ: ");
							} while (spz.size() != 7);
							sheet->writeStr(i, 5, spz.c_str());
						}
					}
					else {
						do {
							spz = inputwstrparms(L"Zadejte SPZ: ");
						} while (spz.size() != 7);
						sheet->writeStr(i, 5, spz.c_str());
					}

					do {
						input = inputwstrparms(L"�idi� nebo spolujezdec? (r/s): ");
					} while (input != L"r" && input != L"s");

					if (input == L"r") {
						input = L"�";
					}
					else {
						input = L"S";
					}

					sheet->writeStr(i, 6, input.c_str());

				}
				else if (input == L"s") {
					activity = L"Swap ";
					target = true;
				}
				else if (input == L"p") {
					activity = L"P��prava ";
				}
				else if (input == L"z") {
					activity = L"Za�i�t�n� ";
				}
				else if (input == L"u") {
					activity = L"Survey ";
				}
				else if (input == L"t") {
					activity = L"Tickets ";
				}
				else if (input == L"d") {
					activity == L"Dokumentace";
				}
				else if (input == L"a") {
					activity == L"Administrativa";
				}

				if (target) {
					do {
						input = inputwstrparms(L"Zadejte c�l cesty/site: ");
					} while (input == L"");
					activity += input;
					sheet->writeStr(i, 7, activity.c_str());
				}
				do {
					input = inputwstrparms(L"Chcete dal�� z�znam do tohoto dne? (a/n): ");
				} while (input != L"a" && input != L"n");

				if (input == L"n") {
					i = 22 + next;
				}
			}
		}
	}
}

std::wstring create_xls::convert_time(unsigned int time) {
	std::wstring string_time{ std::to_wstring(time) };
	if (string_time.size() == 4) {
		string_time.insert(2, L":");
	}
	else if (string_time.size() == 3) {
		string_time.insert(1, L":");
	}
	else if (string_time.size() == 2) {
		string_time.insert(1, L"0:");
	}
	return string_time;
}
