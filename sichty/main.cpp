#include "stdafx.h"
#include <iostream>
#include <fcntl.h>
#include <io.h>
#include "main_functions.h"

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

