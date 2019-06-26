/*
 * Pulseproject.cpp
 *
 *  Created on: Oct 20, 2012
 *      Author: lherrera
 */

#include "Pulseproject.h"
#include "BookMark.h"
#include <stdio.h>
#include <string.h>

#include <list>
using namespace std;

#ifdef _MSC_VER
#include <crtdbg.h>
#endif


void Pulse_project::create_workspace_byPulse	(void)
{
	unsigned_long				Index;
	unsigned_long				IndexFilePwdList;
	unsigned_long				IndexFilePwd;
	unsigned_long				IndexPwd;
	unsigned_long				IndexErrorFile;
	unsigned_long				ul_SheetCount;
	unsigned_long				ul_PulseCount;
	unsigned_long				ul_QtyIndex		= 0;
	unsigned_long				ul_QtySheet		= 0;
	unsigned_long				ul_QtySpread	= 0;
	unsigned_long				ul_QtyPjt		= 0;
	__PwdIndex*					pPwdIndIni;
	__PwdIndex*					pPwdIndEnd;
	__File_Pwd_List_St*			pFilePwdList;
	__WorkSheetBounds*			pWorkSheetBounds;
	__SpreadFile*				pNewSpreadFile;
	list<__SpreadFile*>			SpreadFileList;
	list<__WorkSheetBounds*>	SpreadSheetList;
	list<__PwdIndex*>			ListIndex;
	bool						getSheetInfo	= true;
	bool						getSpreadInfo	= true;
	bool						getPjtInfo		= true;
	__Ptd_St					stSheetPtd;
	__Ptd_St					stSpreadPtd;
	__Ptd_St					stPjtPtd;
	__Ptd_St					stSheetPtdEnd;
	__Ptd_St					stSpreadPtdEnd;
	__Ptd_St					stPjtPtdEnd;
	IndexPwd					= 0;
	IndexFilePwd				= 0;
	IndexFilePwdList			= 0;
	IndexErrorFile				= 0;
	
	//_CrtDumpMemoryLeaks();
	pFilePwdList				= new_FilePwdListSt(IndexFilePwdList);
	Index						= 0;
	ul_PulseCount				= mProject.workSheetConfiguration.PulseQtyCriteria;
	ul_SheetCount				= 0;
	LoadErrPntLst(IndexErrorFile);
	m_cBookMark.SetPoiterList(&mProject.mErrPntList);
	//Init_PwdIndex(&pPwdIndIni);
	do 
	{
		// Agregar Index Ini
//		_CrtDumpMemoryLeaks();
		pPwdIndIni			= new __PwdIndex;
		Add_Index_Ini(&ListIndex,pPwdIndIni,IndexFilePwdList,IndexFilePwd,IndexPwd);
		_AddInfoTopPwd(pPwdIndIni,pFilePwdList);
		_SaveInfo(&stSheetPtd,&pPwdIndIni->st_Ptd,&getSheetInfo);
		_SaveInfo(&stSpreadPtd,&pPwdIndIni->st_Ptd,&getSpreadInfo);
		_SaveInfo(&stPjtPtd,&pPwdIndIni->st_Ptd,&getPjtInfo);
		if (getSheetInfo==true)
		{
			stSheetPtd		= pPwdIndIni->st_Ptd;
		}
		// Agregar Index End
		pPwdIndEnd			= new __PwdIndex;
		ul_QtyIndex			= Add_Pulses_To_Index(pPwdIndEnd,
								pFilePwdList,
								&IndexFilePwdList,
								&IndexFilePwd,
								&IndexPwd,
								&ul_PulseCount);
		_AddInfoTopPwd(pPwdIndEnd,pFilePwdList);
		m_cBookMark.UpdateBookMark(pPwdIndEnd,&stSheetPtd);
		if (m_cBookMark.GetPointOutEnd()==true)
		{  // Save Error File and Load New Error File
			SaveErrPntLst(IndexErrorFile++);
			DestroyErrPntLst();
			LoadErrPntLst(IndexErrorFile);
			m_cBookMark.SetPoiterList(&mProject.mErrPntList);
		}
		ListIndex.push_back(pPwdIndEnd);
		pPwdIndIni->ul_PulseCount	= ul_QtyIndex;
		pPwdIndEnd->ul_PulseCount	= ul_QtyIndex;
		ul_QtySheet			+= ul_QtyIndex;
		
		// Verificar si End of Sheet
		if (ul_PulseCount)
		{
			Next_FilePwdSt(pFilePwdList,&IndexFilePwdList,&IndexFilePwd,&IndexPwd);
			pFilePwdList = mProject.pFilePwdListSt;
		} 
		else
		{
// 			_CrtDumpMemoryLeaks();
			_SaveInfo(&stSheetPtdEnd,&pPwdIndEnd->st_Ptd);
			_SaveInfo(&stSpreadPtdEnd,&pPwdIndEnd->st_Ptd);
			_SaveInfo(&stPjtPtdEnd,&pPwdIndEnd->st_Ptd);
			pWorkSheetBounds	= new_work_sheet_bounds(&ListIndex);
// 			_CrtDumpMemoryLeaks();
			pWorkSheetBounds->ul_PulseCount	= ul_QtySheet;
			ul_QtySpread		+= ul_QtySheet;
			ul_QtySheet			= 0;
			_SaveInfo(&pWorkSheetBounds->stPtdIni,&stSheetPtd);
			_SaveInfo(&pWorkSheetBounds->stPtdEnd,&stSheetPtdEnd);
			SpreadSheetList.push_back(pWorkSheetBounds);
			getSheetInfo		= true;
// 			_CrtDumpMemoryLeaks();
			if (SpreadSheetList.size() >= mProject.workSheetConfiguration.workSheetsPerXlsCount)
			{
				pNewSpreadFile	= new_SpreadFile(&SpreadSheetList);
// 				_CrtDumpMemoryLeaks();
				pNewSpreadFile->us_PulseCount	= ul_QtySpread;
				ul_QtyPjt		+= ul_QtySpread;
				ul_QtySpread	= 0;
				_SaveInfo(&pNewSpreadFile->stPtdIni,&stSpreadPtd);
				_SaveInfo(&pNewSpreadFile->stPtdEnd,&stSpreadPtdEnd);
				SpreadFileList.push_back(pNewSpreadFile);
				getSpreadInfo	= true;
			}
			Next_Pwd(pFilePwdList,&IndexFilePwdList,&IndexFilePwd,&IndexPwd);
			pFilePwdList = mProject.pFilePwdListSt;
			ul_PulseCount		= mProject.workSheetConfiguration.PulseQtyCriteria;
		}
	}
	while(IndexFilePwdList < mProject.FilePwdList_Count);
	//_CrtDumpMemoryLeaks();
	
	//_CrtDumpMemoryLeaks();
	if (ListIndex.size())
	{
		_SaveInfo(&stSheetPtdEnd,&pPwdIndEnd->st_Ptd);
		_SaveInfo(&stSpreadPtdEnd,&pPwdIndEnd->st_Ptd);
		_SaveInfo(&stPjtPtdEnd,&pPwdIndEnd->st_Ptd);
		pWorkSheetBounds	= new_work_sheet_bounds(&ListIndex);
		pWorkSheetBounds->ul_PulseCount	= ul_QtySheet;
		ul_QtySpread		+= ul_QtySheet;
		_SaveInfo(&pWorkSheetBounds->stPtdIni,&stSheetPtd);
		_SaveInfo(&pWorkSheetBounds->stPtdEnd,&stSheetPtdEnd);
		SpreadSheetList.push_back(pWorkSheetBounds);
	}		
	if (SpreadSheetList.size())
	{
		pNewSpreadFile		= new_SpreadFile(&SpreadSheetList);
		pNewSpreadFile->us_PulseCount	= ul_QtySpread;
		ul_QtyPjt			+= ul_QtySpread;
		_SaveInfo(&pNewSpreadFile->stPtdIni,&stSpreadPtd);
		_SaveInfo(&pNewSpreadFile->stPtdEnd,&stSpreadPtdEnd);
		
		SpreadFileList.push_back(pNewSpreadFile);
	}
	_DestroyFilePwdList(mProject.pFilePwdListSt);
	mProject.pFilePwdListSt	= 0;
	mProject.pProjectFile	= new_SpreadFileList(&SpreadFileList);
	_SaveInfo(&mProject.pProjectFile->stPtdIni,&stPjtPtd);
	_SaveInfo(&mProject.pProjectFile->stPtdEnd,&stPjtPtdEnd);
	mProject.pProjectFile->ul_PulseCount = ul_QtyPjt;
	//_CrtDumpMemoryLeaks();
}

void Pulse_project::_SaveInfo	(__Ptd_St* pPtd_Dst,__Ptd_St* pPtd_Src,bool* pFlag)
{
	if (pFlag==NULL || *pFlag==true)
	{
		if(pFlag!=NULL)
		{
			*pFlag		= false;
		}
		*pPtd_Dst	= *pPtd_Src;
	}
}
