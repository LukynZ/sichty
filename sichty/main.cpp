

#include "stdafx.h"
#include "libxl.h"
#include <iostream>
#include <string>
#include <sstream>
#include <fcntl.h>
#include <io.h>
#include <array>

using namespace std;
bool entry();
void start();
void continuing();
int inputparms(wstring);
wstring inputwstrparms(wstring);
class create_xls;

#include "functions.cpp"
#include "main_functions.cpp"

int main() {

	 _setmode(_fileno(stdout), _O_WTEXT);
	 _setmode(_fileno(stdin), _O_WTEXT);
	
	bool beginning;
	beginning = entry();

	if (beginning == true) {
		start();
	}
	else {
		continuing();
	}

    return 0;
}

