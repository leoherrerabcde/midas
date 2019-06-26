#include "gralFunctions.h"
#include <stdio.h>
#include <string.h>

char*	gf_strcpy(char* Str)
{
	char*	pNew = new char[strlen(Str)];
	strcpy(pNew,Str);
	return pNew;
}
