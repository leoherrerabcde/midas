/*
 * cFileName.cpp
 *
 *  Created on: Oct 5, 2012
 *      Author: lherrera
 */

#include "cFileName.h"
#include <stdio.h>
#include <string.h>

cFileName::cFileName() {
	// TODO Auto-generated constructor stub
	mName.FileName		= NULL;
	mName.LengthString	= 0;
}

cFileName::~cFileName() {
	// TODO Auto-generated destructor stub
	if (mName.FileName != NULL)
	{
		delete [] mName.FileName;
		mName.FileName = NULL;
		mName.LengthString = 0;
	}
}

cFileName::cFileName(char *FileName)
{
	mName.LengthString	= strlen(FileName);
	mName.FileName 		= new char[mName.LengthString+1];
	strcpy(mName.FileName, FileName);
}

cFileName::cFileName(char *FileName,short FileNameLength)
{
	mName.LengthString	= FileNameLength;
	mName.FileName 		= new char[FileNameLength+1];
	strcpy(mName.FileName, FileName);
}

void cFileName::SetFileName(char *FileName)
{
	if (mName.FileName != NULL)
	{
		delete [] mName.FileName;
		mName.FileName = NULL;
		mName.LengthString = 0;
	}
	mName.LengthString	= strlen(FileName);
	mName.FileName 		= new char[mName.LengthString+1];
	strcpy(mName.FileName, FileName);
}

void cFileName::SetFileName(char *FileName,short FileNameLength)
{
	if (mName.FileName != NULL)
	{
		delete [] mName.FileName;
		mName.FileName = NULL;
		mName.LengthString = 0;
	}
	mName.LengthString	= FileNameLength;
	mName.FileName 		= new char[FileNameLength+1];
	strcpy(mName.FileName, FileName);
}



