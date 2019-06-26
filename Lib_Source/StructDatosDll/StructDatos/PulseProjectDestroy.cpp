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

template <class T> void	Pulse_project::__DestroyList(list<T*>*pList)
{
	list<T*>::iterator	it;
	for (it=pList->begin();it!=pList->end();it++)
	{
		_Destroy(*it);
	}
	pList->clear();
}

void Pulse_project::DestroyListIndex	(list<__PwdIndex*>* pListIndex)
{
	list<__PwdIndex*>::iterator	it;
	for (it=pListIndex->begin();it!=pListIndex->end();it++)
	{
		delete *it;
	}
	pListIndex->clear();
}

void Pulse_project::DestroyListWorkSheetBounds	(list<__WorkSheetBounds*>* pWorkSheetBounds)
{
	list<__WorkSheetBounds*>::iterator	it;
	__WorkSheetBounds*	p;

	for (it=pWorkSheetBounds->begin();it!=pWorkSheetBounds->end();it++)
	{
		p		= *it;
		delete	[]p->p_StartBoundArray;
		delete	p;
	}
	pWorkSheetBounds->clear();
	//__DestroyList(pWorkSheetBounds);
}

void Pulse_project::DestroyListSpreadFile(list<__SpreadFile*>*pSpreadFileList)
{
	list<__SpreadFile*>::iterator	it;
	__SpreadFile*					p;
	
	for(it=pSpreadFileList->begin();it!=pSpreadFileList->end();it++)
	{
		p			= *it;
		DestroyWorkSheetBoundsArray(p->pWorkSheetArray,p->us_WorkSheetCount);
		delete p;
	}
	pSpreadFileList->clear();
}

void Pulse_project::_DestroyFieldWorkSheetBounds(__WorkSheetBounds* pWorkSheetBounds)
{
	__DestroyArray(pWorkSheetBounds->p_StartBoundArray);
	pWorkSheetBounds->p_StartBoundArray	= NULL;
}

void Pulse_project::DestroyWorkSheetBounds(__WorkSheetBounds* pWorkSheetBounds)
{
	_DestroyFieldWorkSheetBounds(pWorkSheetBounds);
	__Destroy(pWorkSheetBounds) ;
	pWorkSheetBounds	= 0;
}

void Pulse_project::DestroyWorkSheetBoundsArray(__WorkSheetBounds* pWorkSheetBounds,long lCount)
{
	long				i;
	__WorkSheetBounds*	pWrk = pWorkSheetBounds;

	for (i=0;i<lCount;i++)
	{
		_DestroyFieldWorkSheetBounds(pWrk++);
	}
	__DestroyArray(pWorkSheetBounds);
}

void Pulse_project::_DestroyFieldSpreadFile(__SpreadFile* pSpreadFile)
{
	long				i;
	__WorkSheetBounds*	pWorkSheetBounds;
	
	pWorkSheetBounds	= pSpreadFile->pWorkSheetArray;
	for (i=0;i<pSpreadFile->us_WorkSheetCount;i++)
	{
		_DestroyField(pWorkSheetBounds++);
	}
	delete []pSpreadFile->pWorkSheetArray;
}

void Pulse_project::DestroySpreadFile(__SpreadFile* pSpreadFile)
{
	long				i;
	__WorkSheetBounds*	pWorkSheetBounds;

	pWorkSheetBounds	= pSpreadFile->pWorkSheetArray;
	for (i=0;i<pSpreadFile->us_WorkSheetCount;i++)
	{
		_DestroyField(pWorkSheetBounds++);
	}
	delete []pSpreadFile->pWorkSheetArray;
	delete pSpreadFile;
}

void Pulse_project::DestroySpreadFileList(__SpreadFileList* pSpreadFileList)
{
	long				i;
	__SpreadFile*		pSpreadFile;

	if (pSpreadFileList!=NULL)
	{
		pSpreadFile			= pSpreadFileList->pSpreadFileArray;
		for (i=0;i<pSpreadFileList->us_SpreadFileCount;i++)
		{
			_DestroyFieldSpreadFile(pSpreadFile++);
		}
		delete []pSpreadFileList->pSpreadFileArray;
		delete pSpreadFileList;
	}
}

void Pulse_project::__Destroy(void* p)
{
	if (p!=NULL)
	{
		delete p;
	}
}

void Pulse_project::__DestroyArray(void* p)
{
	if (p!=NULL)
	{
		delete []p;
	}
}

void Pulse_project::_DestroyFilePwdList(__File_Pwd_List_St* pFilePwdLstSt)
{
	unsigned_long	i;

	if (pFilePwdLstSt!=NULL)
	{
		_Destroy(pFilePwdLstSt->p_PathName);
		_Destroy(pFilePwdLstSt->p_st_FileList);
		for(i=0;i<pFilePwdLstSt->us_ListCount;i++)
		{
			_DestroyField(pFilePwdLstSt->p_st_FilePwdArray+i);
		}
		__DestroyArray(pFilePwdLstSt->p_st_FilePwdArray);
		delete pFilePwdLstSt;
	}
}

void Pulse_project::_DestroyFieldFilePwdSt(__File_Pwd_St* pFilePwdSt)
{
	if (pFilePwdSt!=NULL)
	{
		__DestroyArray(pFilePwdSt->p_ch_FileName);
		__DestroyArray(pFilePwdSt->p_st_Pwd);
		__DestroyArray(pFilePwdSt->p_st_Pwd_NewFields);
		pFilePwdSt->p_ch_FileName	= NULL;
		pFilePwdSt->p_st_Pwd		= NULL;
		pFilePwdSt->p_st_Pwd_NewFields= NULL;
	}
}

void Pulse_project::_DestroyFilePwdSt(__File_Pwd_St* pFilePwdSt)
{
	if (pFilePwdSt!=NULL)
	{
		__DestroyArray(pFilePwdSt->p_ch_FileName);
		__DestroyArray(pFilePwdSt->p_st_Pwd);
		__DestroyArray(pFilePwdSt->p_st_Pwd_NewFields);
		delete pFilePwdSt;
	}
}

void Pulse_project::_DestroyFileList(__File_List_St* pFileListSt)
{
	if(pFileListSt!=NULL) {
		__DestroyArray(pFileListSt->p_ch_FileList);
		__DestroyArray(pFileListSt->p_ch_NamesList);
		__DestroyArray(pFileListSt->p_ch_Path);
		__DestroyArray(pFileListSt->p_us_NamesLenList);
		delete pFileListSt;
	}
}

void Pulse_project::DestroyWorkSpace	(void)
{
	_Destroy(mProject.pProjectFile);
	mProject.pProjectFile		= NULL;
}

void Pulse_project::DestroyAll	(void)
{
	_Destroy(mProject.pFileListSt);
	_Destroy(mProject.pFilePwdListSt);
	_Destroy(mProject.pFilePwdSt);
	_Destroy(mProject.pProjectFile);
	DestroyErrPntLst();
	mProject.pFileListSt		= NULL;
	mProject.pFilePwdListSt		= NULL;
	mProject.pFilePwdSt			= NULL;
	mProject.pProjectFile		= NULL;
}

void Pulse_project::DestroyErrPntLst(void)
{
	if (mProject.mErrPntList.PointerArray!=NULL)
	{
		delete []mProject.mErrPntList.PointerArray;
		mProject.mErrPntList.PointerArray	= NULL;
		mProject.mErrPntList.Count			= 0;
	}
}

void Pulse_project::UnSetErrPntLst(void)
{
	if (mProject.mErrPntList.PointerArray!=NULL)
	{
		mProject.mErrPntList.PointerArray	= NULL;
		mProject.mErrPntList.Count			= 0;
	}
}