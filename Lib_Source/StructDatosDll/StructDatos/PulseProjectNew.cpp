/*
 * Pulseproject.cpp
 *
 *  Created on: Oct 20, 2012
 *      Author: lherrera
 */

#include "Pulseproject.h"
#include <stdio.h>
#include <string.h>
#ifdef _MSC_VER
#include <crtdbg.h>
#endif


#include <list>
using namespace std;

__WorkSheetBounds* Pulse_project::new_work_sheet_bounds(list<__PwdIndex*>* pListIndex)
{
	__WorkSheetBounds*	pWorkSheetBounds;
	list<__PwdIndex*>::iterator	it;
	__PwdIndex*			pPwdIndex;
	
	pWorkSheetBounds	= new __WorkSheetBounds;
	pPwdIndex			= new struct __pwd_index[pListIndex->size()];
	pWorkSheetBounds->ul_BoundsCount	= pListIndex->size();
	pWorkSheetBounds->p_StartBoundArray	= pPwdIndex;
	for(it=pListIndex->begin();it!=pListIndex->end();it++)
	{
		memcpy(pPwdIndex++,*it,sizeof(__PwdIndex));
	}
	
	DestroyListIndex(pListIndex);
	return pWorkSheetBounds;
}

__SpreadFile* Pulse_project::new_SpreadFile(list<__WorkSheetBounds*>* pWorkSheetBoundsList)
{
	__SpreadFile*			pSpreadFile;
	list<__WorkSheetBounds*>::iterator	it;
	__WorkSheetBounds*		pWorkSheetBounds;
	__WorkSheetBounds*		pWrkTmp;
	//__PwdIndex*				pPwdIndexArray;
	
	pSpreadFile				= new __SpreadFile;
	pWorkSheetBounds		= new __WorkSheetBounds[pWorkSheetBoundsList->size()];
	
	pSpreadFile->us_WorkSheetCount	= pWorkSheetBoundsList->size();
	pSpreadFile->pWorkSheetArray	= pWorkSheetBounds;
	for(it=pWorkSheetBoundsList->begin();it!=pWorkSheetBoundsList->end();it++)
	{
		pWrkTmp						= *it;
		memcpy(pWorkSheetBounds++,pWrkTmp,sizeof(__WorkSheetBounds));
		pWrkTmp->p_StartBoundArray	= new __PwdIndex[2];
	}
	
	DestroyListWorkSheetBounds(pWorkSheetBoundsList);
	return pSpreadFile;
}

__SpreadFileList* Pulse_project::new_SpreadFileList(list<__SpreadFile*>*pSpreadFileList)
{
	list<__SpreadFile*>::iterator it;
	__SpreadFileList* pNewSpreadFileList;
	__SpreadFile*		pSpreadFile;
	__SpreadFile*		pTmp;
	
	pNewSpreadFileList	= new __SpreadFileList;
	pSpreadFile			= new __SpreadFile[pSpreadFileList->size()];
	pNewSpreadFileList->us_SpreadFileCount	= pSpreadFileList->size();
	pNewSpreadFileList->pSpreadFileArray	= pSpreadFile;
	for(it=pSpreadFileList->begin();it!=pSpreadFileList->end();it++)
	{
		pTmp					= *it;
		memcpy(pSpreadFile++,pTmp,sizeof(__SpreadFile));
		pTmp->pWorkSheetArray	= new __WorkSheetBounds[1];
		pTmp->us_WorkSheetCount	= 1;
		pTmp->pWorkSheetArray->p_StartBoundArray=NULL;
	}
	DestroyListSpreadFile(pSpreadFileList);
	return pNewSpreadFileList;
}

__File_Pwd_St* Pulse_project::new_FilePwdSt(unsigned_long ulQty)
{
	__File_Pwd_St*	pNewFilePwdSt	= new __File_Pwd_St;

	pNewFilePwdSt->l_Pulse_Count	= ulQty;
	pNewFilePwdSt->p_st_Pwd			= new __Pwd_St[ulQty];
	pNewFilePwdSt->p_st_Pwd_NewFields	= new __Pwd_NF_St[ulQty];
	pNewFilePwdSt->p_ch_FileName	= NULL;

	return pNewFilePwdSt;
}

void Pulse_project::_cpyIndexData(__Pwd_St* pPwd,__Pwd_NF_St* pPwdNf,__PwdIndex* pIndex)
{
	__File_Pwd_List_St*		pFilePwdList;
	__File_Pwd_St*			pFilePwdSt;
	__Pwd_NF_St*			pPwdNFSrc;
	__Pwd_St*				pPwdSrc;

	pFilePwdList		= mProject.pFilePwdListSt;
	if (pIndex->ul_IndexWorkSpace!=mProject.IndexFilePwdLstSt || pFilePwdList==NULL)
	{
		//_CrtDumpMemoryLeaks();
		if (pFilePwdList==NULL)
		{
			pFilePwdList	= new_FilePwdListSt();
			mProject.pFilePwdListSt			= pFilePwdList;
			//_CrtDumpMemoryLeaks();
		}
		set_filePwdListSt(pFilePwdList,pIndex->ul_IndexWorkSpace);
		//_CrtDumpMemoryLeaks();
		read_file(pIndex->ul_IndexWorkSpace);
		//_CrtDumpMemoryLeaks();
	} 
	pFilePwdSt			= pFilePwdList->p_st_FilePwdArray + pIndex->us_IndexFilePwd;
	pPwdSrc				= pFilePwdSt->p_st_Pwd + pIndex->us_IndexPulse;
	pPwdNFSrc			= pFilePwdSt->p_st_Pwd_NewFields + pIndex->us_IndexPulse;
	memcpy(pPwd,pPwdSrc,sizeof(__Pwd_St)*pIndex->ul_PulseCount);
	memcpy(pPwdNf,pPwdNFSrc,sizeof(__Pwd_NF_St)*pIndex->ul_PulseCount);
}

void Pulse_project::_IndexToPointers(__PwdIndex* pIndex,
									 __File_Pwd_St **pFilePwd,
									 __Pwd_NF_St** pPwdNf,
									 __Pwd_St** pPwd)
{
	__File_Pwd_List_St*		pFilePwdLst = mProject.pFilePwdListSt;
	
	if (pFilePwdLst!=NULL)
	{
		if (pFilePwdLst->p_st_FilePwdArray!=NULL)
		{
			*pFilePwd			= pFilePwdLst->p_st_FilePwdArray + pIndex->us_IndexFilePwd;
			*pPwd				= (*pFilePwd)->p_st_Pwd + pIndex->us_IndexPulse;
			*pPwdNf				= (*pFilePwd)->p_st_Pwd_NewFields + pIndex->us_IndexPulse;
		} 
		else
		{
			*pFilePwd			= NULL;
			*pPwd				= NULL;
			*pPwdNf				= NULL;
		}
	} 
	else
	{
		*pFilePwd			= NULL;
		*pPwd				= NULL;
		*pPwdNf				= NULL;
	}
}

void Pulse_project::_AddInfoTopPwd (__PwdIndex* pIndex,__File_Pwd_List_St* pFilePwdList)
{
	__File_Pwd_St*		pFilePwdSt;
	__Pwd_NF_St*		pPwdNf;
	__Pwd_St*			pPwd;
	__Ptd_St*			pPtd;

	_IndexToPointers(pIndex,&pFilePwdSt,&pPwdNf,&pPwd);
	if (pFilePwdSt!=NULL && pPwd!=NULL && pPwdNf!=NULL)
	{
		pPtd				= &(pIndex->st_Ptd);
		pPtd->d_D_Toa_ms	= pPwdNf->d_post_d_Toa_us;
		pPtd->d_FileTime_s	= pPwdNf->d_Abs_Toa_s;
		pPtd->d_postDtoa_us	= pPwdNf->d_post_d_Toa_us;
		pPtd->d_Time_ms		= pPwdNf->d_Rel_Toa_ms;
		pPtd->ul_Toa		= pPwd->ul_Toa;
		pPtd->ul_Index		= pPwdNf->ul_Index;
	} 
	else
	{
		pFilePwdSt	= NULL;
	}
}


bool Pulse_project::Create_SpreadFile(__File_Pwd_List_St* pFilePwdLstSt,long IndexSpread)
{
	if (pFilePwdLstSt==NULL)
	{
		return false;
	}
	if (pFilePwdLstSt->us_ListCount != mProject.pProjectFile->pSpreadFileArray[IndexSpread].us_WorkSheetCount)
	{
	}
	return false;
}


__File_Pwd_List_St* Pulse_project::Create_SpreadFile(long IndexSpread)
{
	__File_Pwd_List_St* pFilePwdListSt;
	long				i;

	//_CrtDumpMemoryLeaks();
	pFilePwdListSt					= new __File_Pwd_List_St;
	//_CrtDumpMemoryLeaks();
	pFilePwdListSt->us_ListCount	= mProject.pProjectFile->pSpreadFileArray[IndexSpread].us_WorkSheetCount;
	pFilePwdListSt->p_PathName		= NULL;
	pFilePwdListSt->p_st_FileList	= NULL;
	pFilePwdListSt->p_st_FilePwdArray	= new __File_Pwd_St[pFilePwdListSt->us_ListCount];
	//_CrtDumpMemoryLeaks();

	for (i=0;i<pFilePwdListSt->us_ListCount;i++)
	{
		_Get_Sheet(pFilePwdListSt->p_st_FilePwdArray+i,IndexSpread,i);
		//_CrtDumpMemoryLeaks();
// 		if (pFilePwdListSt->p_st_FilePwdArray->l_Pulse_Count!=3413)
// 		{
// 			pFilePwdListSt->p_st_FilePwdArray->l_Pulse_Count=0;
// 		}
	}
	return pFilePwdListSt;
}

__File_Pwd_List_St* Pulse_project::new_FilePwdListSt(void)
{
	__File_Pwd_List_St* pFilePwdList = new __File_Pwd_List_St;
	pFilePwdList->p_st_FilePwdArray	= 0;
	pFilePwdList->ul_PulseCount		= 0;
	pFilePwdList->us_ListCount		= 0;
	pFilePwdList->p_PathName		= NULL;
	pFilePwdList->p_st_FileList		= NULL;
	pFilePwdList->p_st_FilePwdArray	= NULL;
	
	return pFilePwdList;
}

