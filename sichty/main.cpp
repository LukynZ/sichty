

#include "stdafx.h"
#include "libxl.h"
#include <iostream>
#include <string>
#include <sstream>
#include <fcntl.h>
#include <io.h>
#include <array>
#include "create_xls.h"
#include "main_functions.cpp"
#include "create_xls.cpp"

using namespace std;

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

