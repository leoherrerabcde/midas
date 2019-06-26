/*
 * Pulseproject.cpp
 *
 *  Created on: Oct 5, 2012
 *      Author: lherrera
 */

#include "Pulseproject.h"
#include <stdio.h>
#include <string.h>

#include <list>
using namespace std;

unsigned_long Pulse_project::Get_FileMaxPulses(__File_List_St* pFileLstSt,
											   unsigned_long lvFileLstCount)
{
	char**			pFileName;
	char			FileName[260];
	char*			pDst;
	unsigned_long	lvMax		=0;
	unsigned_long	lvPulseCount;
	unsigned_long	i,j;
	
	strcpy(FileName,pFileLstSt->p_ch_Path);
	strcat(FileName,"\\");
	pDst			= &FileName[strlen(FileName)];
		
	for (j=0;j<lvFileLstCount;j++)
	{
		pFileName	= pFileLstSt->p_ch_FileList;
		for (i=0;i<pFileLstSt->us_ListCount;i++)
		{
			strcpy(pDst,*pFileName);
			lvPulseCount	= Get_FilePulses(FileName);
			if(lvMax<lvPulseCount)
			{
				lvMax	= lvPulseCount;
			}
			pFileName	++;
		}
		pFileLstSt++;
	}
	return lvMax;
}

unsigned_long Pulse_project::Get_FilePulses(char* FileName)
{
	FILE			*pFile;
	long			l_FileSize;

	pFile               = fopen(FileName,"rb");
	if (pFile == NULL)
	{
		return 0;
	}
	fseek( pFile, 0, SEEK_END );
	l_FileSize	=  ftell( pFile );
	fclose(pFile);
	return l_FileSize / sizeof(__Pwd_St);
}