#ifndef CREATE_XLS_H
#define CREATE_XLS_H

class create_xls {
	
	libxl::Book * book;
	libxl::Sheet * sheet;
	libxl::Format * format;

public:

	create_xls(short month, int year, std::wstring name);
	~create_xls();
	void set_formating(unsigned short next);
	void align(libxl::AlignH styleH, libxl::AlignV styleV, libxl::Format * format);
	void border(libxl::BorderStyle styleT, libxl::BorderStyle styleB, libxl::BorderStyle styleL, libxl::BorderStyle styleR, libxl::Format * format);
	void setformat(unsigned short row, unsigned short col, libxl::Format * format, unsigned short next);
	void basefont(libxl::Format * format);
	void datafont(libxl::Format * format);
	void boldfont(libxl::Format * format);
	void layout(unsigned short next);
	void fill_main_text(unsigned short next, std::wstring name, unsigned short day, unsigned short month, unsigned int year);
	bool check_fill_data(std::wstring date, unsigned short day);
	void fill_data(unsigned short next);
	std::wstring convert_time(unsigned int time);
};
#endif // CREATE_XLS_H

