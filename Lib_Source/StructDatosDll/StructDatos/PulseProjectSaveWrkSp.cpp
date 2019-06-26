/*
 * Pulseproject.cpp
 *
 *  Created on: Oct 20, 2012
 *      Author: lherrera
 */

#include "Pulseproject.h"
#include <stdio.h>
#include <string.h>

#include <list>
using namespace std;

bool Pulse_project::SaveWorkSpace(void)
{
	FILE*				pFile;
	char				FileName[260];
	__SpreadFileList*	pSpreadFileListSt;
	__SpreadFile*		pSpreadFileArray;
	__WorkSheetBounds*	pWrkSheetBounds;
	__PwdIndex*			pIndexBound;
	long				Index,n;
	
	sprintf(FileName,"%s\\WrkSpc.bin",mProject.workSpacePath->FileName);
	pFile		= fopen(FileName,"wb");
	if (pFile==NULL)
	{
		return false;
	}
	
	
	pSpreadFileListSt	= mProject.pProjectFile;
	
	if(write_in_file(pSpreadFileListSt,sizeof(__SpreadFileList),1,pFile)==false)
	{
		return false;
	}

	
	pSpreadFileArray	= pSpreadFileListSt->pSpreadFileArray;
	
	if(write_in_file(pSpreadFileArray,sizeof(__SpreadFile),pSpreadFileListSt->us_SpreadFileCount,pFile)==false)
	{
		return false;
	}
	
	
	for(Index=0;Index<pSpreadFileListSt->us_SpreadFileCount;Index++)
	{
		pWrkSheetBounds	= pSpreadFileArray->pWorkSheetArray;
		if(write_in_file(pWrkSheetBounds,sizeof(__WorkSheetBounds),pSpreadFileArray->us_WorkSheetCount,pFile)==false)
		{
			return false;
		}

		for (n=0;n<pSpreadFileArray->us_WorkSheetCount;n++)
		{
			pIndexBound	= pWrkSheetBounds->p_StartBoundArray;
			if(write_in_file(pIndexBound,sizeof(__PwdIndex),pWrkSheetBounds->ul_BoundsCount,pFile)==false)
			{
				return false;
			}
			pWrkSheetBounds ++;
		}
		pSpreadFileArray++;
	}
	fclose(pFile);
	return true;
}


bool Pulse_project::LoadWorkSpace(void)
{
	FILE*				pFile;
	char				FileName[260];
	__SpreadFileList*	pSpreadFileListSt;
	__SpreadFile*		pSpreadFileArray;
	__WorkSheetBounds*	pWrkSheetBounds;
	__PwdIndex*			pIndexBound;
	long				Index,n;
	
	sprintf(FileName,"%s\\WrkSpc.bin",mProject.workSpacePath->FileName);
	pFile		= fopen(FileName,"rb");
	if (pFile==NULL)
	{
		return false;
	}
	
	pSpreadFileListSt	= new __SpreadFileList;
	mProject.pProjectFile	= pSpreadFileListSt;
	
	if(read_in_file(pSpreadFileListSt,sizeof(__SpreadFileList),1,pFile)==false)
	{
		return false;
	}
	
	pSpreadFileArray	= new __SpreadFile[pSpreadFileListSt->us_SpreadFileCount];
	pSpreadFileListSt->pSpreadFileArray	= pSpreadFileArray;
	
	if(read_in_file(pSpreadFileArray,sizeof(__SpreadFile),pSpreadFileListSt->us_SpreadFileCount,pFile)==false)
	{
		return false;
	}	
	
	for(Index=0;Index<pSpreadFileListSt->us_SpreadFileCount;Index++)
	{
		pWrkSheetBounds	= new __WorkSheetBounds[pSpreadFileArray->us_WorkSheetCount];
		pSpreadFileArray->pWorkSheetArray	= pWrkSheetBounds;
		if(read_in_file(pWrkSheetBounds,sizeof(__WorkSheetBounds),pSpreadFileArray->us_WorkSheetCount,pFile)==false)
		{
			return false;
		}
		
		for (n=0;n<pSpreadFileArray->us_WorkSheetCount;n++)
		{
			pIndexBound	= new __PwdIndex[pWrkSheetBounds->ul_BoundsCount];
			pWrkSheetBounds->p_StartBoundArray	= pIndexBound;
			if(read_in_file(pIndexBound,sizeof(__PwdIndex),pWrkSheetBounds->ul_BoundsCount,pFile)==false)
			{
				return false;
			}
			pWrkSheetBounds ++;
		}
		pSpreadFileArray++;
	}

	fclose(pFile);
	return true;
}


bool Pulse_project::SaveColumnFormat(void)
{
	//memcpy(mProject.workSheetConfiguration.ColumnOrder,ColumnOrder,PWD_FIELD_COUNT*sizeof(long));
	//memcpy(mProject.workSheetConfiguration.ColumnEnable,ColumnEnable,PWD_FIELD_COUNT*sizeof(long));
	return true;
}
