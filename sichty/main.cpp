// sichty.cpp: Definuje vstupní bod pro konzolovou aplikaci.
//

#include "stdafx.h"
#include "libxl.h"
#include <iostream>
#include <string>
#include <sstream>
#include <fcntl.h>
#include <io.h>
#include "main_functions.cpp"
#include "functions.cpp"

using namespace std;
bool entry();
void start();
void continuing();

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

