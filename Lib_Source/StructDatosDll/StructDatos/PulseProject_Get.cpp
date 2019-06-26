/*
 * Pulseproject.cpp
 *
 *  Created on: Oct 20, 2012
 *      Author: lherrera
 */

#include "pulseformat.h"
#include "Pulseproject.h"
#include <stdio.h>
#include <string.h>
#ifdef _MSC_VER
#include <crtdbg.h>
#endif

#include <list>
using namespace std;

long Pulse_project::GetPulseCount(long IndexSpread,long IndexSheet)
{
	return mProject.pProjectFile->pSpreadFileArray[IndexSpread].pWorkSheetArray[IndexSheet].ul_PulseCount;
}

__WorkSheetBounds* Pulse_project::_Get_WorkSheetBound(long IndexSpread,long IndexSheet)
{
	__SpreadFile* p		= _Get_SpreadFile(IndexSpread);
	return p->pWorkSheetArray + IndexSpread;
}

__SpreadFile* Pulse_project::_Get_SpreadFile(long IndexSpread)
{
	return mProject.pProjectFile->pSpreadFileArray + IndexSpread;
}

__PwdIndex* Pulse_project::_Get_PwdIndex(__WorkSheetBounds* pWorkSheetBounds,
										 long IndexPulse)
{
	unsigned_long		i,Index;
	__PwdIndex*			pPwdIndexIni;
	__PwdIndex*			pPwdIndexEnd;

	pPwdIndexIni		= pWorkSheetBounds->p_StartBoundArray;
	return pPwdIndexIni;
	pPwdIndexEnd		= pPwdIndexIni + 1;
	for (i=0;i<pWorkSheetBounds->ul_BoundsCount;i+=2)
	{
		if(pPwdIndexEnd->ul_PulseCount>IndexPulse)
		{
			Index		= pPwdIndexIni->us_IndexPulse + IndexPulse;
		}
	}
}

/*__File_Pwd_St* Pulse_project::_Get_FilePwd(long IndexSpread,
										   long IndexSheet,
										   long IndexPulse)
{
	__PwdIndex*	pIndex;
	__WorkSheetBounds* pWorkSheetBounds;
	
	pWorkSheetBounds	= _Get_WorkSheetBound(IndexSpread,IndexSheet);
}

__File_Pwd_St* Pulse_project::_Get_FilePwd(unsigned_long IndFilePwdLst,
										   unsigned_long IndFilePwd)
{
	if (mProject.IndexFilePwdLstSt != IndFilePwdLst)
	{
		
	} 
	else
	{
	}
}

__File_Pwd_St* Pulse_project::_Get_FilePwd(unsigned_long IndFilePwd)
{
	if (mProject.IndexFilePwdSt!=IndFilePwd)
	{
		mProject.IndexFilePwdSt		= IndFilePwd;
		mProject.pFilePwdSt		= mProject.pFilePwdListSt->p_st_FilePwdArray + IndFilePwd;
	}
	return mProject.pFilePwdSt;
}*/

__File_Pwd_List_St* Pulse_project::_Get_FilePwdList(unsigned_long IndFilePwdLst)
{
	if (mProject.IndexFilePwdLstSt != IndFilePwdLst)
	{
		__File_Pwd_List_St*		pFilePwdList	= mProject.pFilePwdListSt;
		if (pFilePwdList==NULL)
		{
			pFilePwdList	= new_FilePwdListSt(IndFilePwdLst);
		}
		else
		{
			//_Destroy(pFilePwdList);
			set_filePwdListSt(pFilePwdList,IndFilePwdLst);
			read_filePwdListSt(IndFilePwdLst);
		}
	}
	return mProject.pFilePwdListSt;
}

void Pulse_project::_Get_Sheet(__File_Pwd_St* pFilePwdSt,long IndexSpread,long IndexSheet)
{
	__SpreadFile*		pSpreadFile;
	__WorkSheetBounds*	pWorkSheetBounds;
	__PwdIndex*			pIndex;
	__Pwd_St*			pPwd;
	__Pwd_NF_St*		pPwdNf;
	unsigned_long		i;

	//_CrtDumpMemoryLeaks();
	pSpreadFile				= mProject.pProjectFile->pSpreadFileArray + IndexSpread;
	pWorkSheetBounds		= pSpreadFile->pWorkSheetArray + IndexSheet;
	
	pFilePwdSt->l_Pulse_Count	= pWorkSheetBounds->ul_PulseCount;
	pFilePwdSt->p_st_Pwd			= new __Pwd_St[pWorkSheetBounds->ul_PulseCount];
	//_CrtDumpMemoryLeaks();
	pFilePwdSt->p_st_Pwd_NewFields	= new __Pwd_NF_St[pWorkSheetBounds->ul_PulseCount];
	//_CrtDumpMemoryLeaks();
	pFilePwdSt->p_ch_FileName	= NULL;
	
	pPwd					= pFilePwdSt->p_st_Pwd;
	pPwdNf					= pFilePwdSt->p_st_Pwd_NewFields;
	pIndex					= pWorkSheetBounds->p_StartBoundArray;
	
	for (i=0;i<pWorkSheetBounds->ul_BoundsCount;i+=2)
	{
		_cpyIndexData(pPwd,pPwdNf,pIndex);
		//_CrtDumpMemoryLeaks();
		pPwd		+= pIndex->ul_PulseCount;
		pPwdNf		+= pIndex->ul_PulseCount;
		pIndex		+= 2;
	}
}

void Pulse_project::_Get_Sheet(long IndexSpread,long IndexSheet)
{
	__Temp_Index*		Tempo	= &mProject.TempIndex;
	__SpreadFile*		pSpreadFile;
	__WorkSheetBounds*	pWorkSheetBounds;
	__File_Pwd_St*		pFilePwdSt;
	__PwdIndex*			pIndex;
	//__PwdIndex*			pEnd;
	__Pwd_St*			pPwd;
	__Pwd_NF_St*		pPwdNf;
	unsigned_long		i;

	if ((Tempo->IndexSpread!=IndexSpread) || (Tempo->IndexSheet!=IndexSheet) || mProject.pFilePwdListSt == NULL)
	{
		Tempo->IndexSheet		= IndexSheet;
		Tempo->IndexSpread		= IndexSpread;
		pSpreadFile				= mProject.pProjectFile->pSpreadFileArray + IndexSpread;
		pWorkSheetBounds		= pSpreadFile->pWorkSheetArray + IndexSheet;
		
		_Destroy(mProject.pFilePwdSt);
		pFilePwdSt				= new_FilePwdSt(pWorkSheetBounds->ul_PulseCount);
		mProject.pFilePwdSt		= pFilePwdSt;

		pPwd					= pFilePwdSt->p_st_Pwd;
		pPwdNf					= pFilePwdSt->p_st_Pwd_NewFields;
		pIndex					= pWorkSheetBounds->p_StartBoundArray;

		for (i=0;i<pWorkSheetBounds->ul_BoundsCount;i+=2)
		{
			_cpyIndexData(pPwd,pPwdNf,pIndex);
			pPwd		+= pIndex->ul_PulseCount;
			pPwdNf		+= pIndex->ul_PulseCount;
			pIndex		+= 2;
		}
	} 
}

void Pulse_project::GetPwd(long IndexSpread,
						   long IndexSheet,
						   long IndexPulse,
						   double* pPwd)
{
	pulse_format	mPulseFormat;
	__File_Pwd_St	*p_st_FilePwd;
	__Temp_Index*	Tempo	= &mProject.TempIndex;
	__File_Pwd_List_St*	pFilePwdLstSt;

	pFilePwdLstSt	= mProject.pFilePwdListSt;
	_Get_Sheet(IndexSpread,IndexSheet);
	p_st_FilePwd	= mProject.pFilePwdSt;
	mPulseFormat.format_pwd(p_st_FilePwd,IndexPulse,pPwd);
}

void Pulse_project::get_FileName(long IndexSpread , long IndexSheet ,char* lvStr)
{
	__Temp_Index*	Tempo	= &mProject.TempIndex;
	__File_Pwd_List_St*	pFilePwdLstSt;
	unsigned_long	Index;
	//__SpreadFile*	pSpreadFile;
	//__WorkSheetBounds* pWorkSheetBounds;
	//__File_Pwd_St	*p_st_FilePwd;
	
	_Get_Sheet(IndexSpread,IndexSheet);
	pFilePwdLstSt	= mProject.pFilePwdListSt;
	Index	= mProject.pProjectFile->pSpreadFileArray[IndexSpread].pWorkSheetArray[IndexSheet].p_StartBoundArray->us_IndexFilePwd;
	//pWorkSheetBounds = pSpreadFile->;
	//= mProject.pFilePwdListSt->p_st_FilePwdArray+Index;
	//Index			= pWorkSheetBounds->p_StartBoundArray->us_IndexFilePwd;
	strcpy(lvStr,pFilePwdLstSt->p_st_FilePwdArray[Index].p_ch_FileName);
}
