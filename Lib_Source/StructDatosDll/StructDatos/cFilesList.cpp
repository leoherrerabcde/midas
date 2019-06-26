/*
 * cFilesList.cpp
 *
 *  Created on: Oct 5, 2012
 *      Author: lherrera
 */

#include "cFilesList.h"
#include <stdio.h>

cFilesList::cFilesList() {
	// TODO Auto-generated constructor stub
	mCount			= 0;
	mFileListHead	= NULL;
}

cFilesList::~cFilesList() {
	// TODO Auto-generated destructor stub
	short 			i;
	__File_List_St*	pFileList;
	char*			pPath	= NULL;

	pFileList		= mFileListHead;
	for (i=0; i<mCount; i++)
	{
		if (pPath != pFileList->p_ch_Path)
		{
			pPath	= pFileList->p_ch_Path;
			delete [] pPath;
		}
		delete [] pFileList->p_ch_NamesList;
		delete [] pFileList->p_ch_FileList;
		delete [] pFileList->p_us_NamesLenList;
		pFileList ++;
	}
	delete []mFileListHead;
}

void cFilesList::SetSize(long Count)
{
	mCount			= Count;
	mFileListHead	= new __File_List_St[Count];
}
