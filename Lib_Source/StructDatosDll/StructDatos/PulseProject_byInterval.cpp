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

void Pulse_project::create_workspace_byInterval	(void)
{
	unsigned_long				Index;
	unsigned_long				IndexFilePwdList;
	unsigned_long				IndexFilePwd;
	unsigned_long				IndexPwd;
	unsigned_long				ul_SheetCount;
	// unsigned_long				ul_PulseCount;
	unsigned_long				ul_QtyIndex		= 0;
	unsigned_long				ul_QtySheet		= 0;
	unsigned_long				ul_QtySpread	= 0;
	unsigned_long				ul_QtyPjt		= 0;
	//double						d_IntervalCriteria;
	double						dTimeNextPulse;
	double						d_TimeEnd_ms;
	//double						d_TimeIni_ms;
	double						d_Time_ms;
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
	bool						nextPulseDone;
	__Ptd_St					stSheetPtd;
	__Ptd_St					stSpreadPtd;
	__Ptd_St					stPjtPtd;
	__Ptd_St					stSheetPtdEnd;
	__Ptd_St					stSpreadPtdEnd;
	__Ptd_St					stPjtPtdEnd;
	long						cntDbg = 0;
	IndexPwd					= 0;
	IndexFilePwd				= 0;
	IndexFilePwdList			= 0;
	
	pFilePwdList				= new_FilePwdListSt(IndexFilePwdList);
	Index						= 0;
	d_TimeEnd_ms				= mProject.workSheetConfiguration.IntervalTimeCriteria;
	ul_SheetCount				= 0;

	m_cBookMark.SetPoiterList(&mProject.mErrPntList);

	do 
	{
		cntDbg ++;
		// Agregar Index Ini
		pPwdIndIni			= new __PwdIndex;
		Add_Index_Ini(&ListIndex,pPwdIndIni,IndexFilePwdList,IndexFilePwd,IndexPwd);
		_AddInfoTopPwd(pPwdIndIni,pFilePwdList);
		_SaveInfo(&stSheetPtd,&pPwdIndIni->st_Ptd,&getSheetInfo);
		_SaveInfo(&stSpreadPtd,&pPwdIndIni->st_Ptd,&getSpreadInfo);
		_SaveInfo(&stPjtPtd,&pPwdIndIni->st_Ptd,&getPjtInfo);

		// Agregar Index End
		pPwdIndEnd			= new __PwdIndex;
		ul_QtyIndex			= Add_Pulses_To_Index(pPwdIndEnd,
								pFilePwdList,
								&IndexFilePwdList,
								&IndexFilePwd,
								&IndexPwd,
								&d_Time_ms,
								d_TimeEnd_ms);
		_AddInfoTopPwd(pPwdIndEnd,pFilePwdList);
		m_cBookMark.UpdateBookMark(pPwdIndEnd,&stSheetPtd);
		ListIndex.push_back(pPwdIndEnd);
		pPwdIndIni->ul_PulseCount	= ul_QtyIndex;
		pPwdIndEnd->ul_PulseCount	= ul_QtyIndex;
		ul_QtySheet			+= ul_QtyIndex;
		nextPulseDone		= false;

		// Verificar si End of Sheet
		if ((d_TimeEnd_ms - d_Time_ms) >= 0.5)
		{
			Next_FilePwdSt(pFilePwdList,&IndexFilePwdList,&IndexFilePwd,&IndexPwd);
			pFilePwdList		= mProject.pFilePwdListSt;
			nextPulseDone		= true;
			dTimeNextPulse		= TimeNextPulse(pFilePwdList,
												IndexFilePwdList,
												IndexFilePwd,
												IndexPwd);
			if ( d_TimeEnd_ms - dTimeNextPulse>= 0.5)
			{
				continue;
			} 
		} 
		_SaveInfo(&stSheetPtdEnd,&pPwdIndEnd->st_Ptd);
		_SaveInfo(&stSpreadPtdEnd,&pPwdIndEnd->st_Ptd);
		_SaveInfo(&stPjtPtdEnd,&pPwdIndEnd->st_Ptd);
		pWorkSheetBounds	= new_work_sheet_bounds(&ListIndex);
		pWorkSheetBounds->ul_PulseCount	= ul_QtySheet;
		ul_QtySpread		+= ul_QtySheet;
		ul_QtySheet			= 0;
		_SaveInfo(&pWorkSheetBounds->stPtdIni,&stSheetPtd);
		_SaveInfo(&pWorkSheetBounds->stPtdEnd,&stSheetPtdEnd);
		SpreadSheetList.push_back(pWorkSheetBounds);
		getSheetInfo		= true;
		if (SpreadSheetList.size() >= mProject.workSheetConfiguration.workSheetsPerXlsCount)
		{
			pNewSpreadFile	= new_SpreadFile(&SpreadSheetList);
			pNewSpreadFile->us_PulseCount	= ul_QtySpread;
			ul_QtyPjt		+= ul_QtySpread;
			ul_QtySpread	= 0;
			_SaveInfo(&pNewSpreadFile->stPtdIni,&stSpreadPtd);
			_SaveInfo(&pNewSpreadFile->stPtdEnd,&stSpreadPtdEnd);
			SpreadFileList.push_back(pNewSpreadFile);
			getSpreadInfo	= true;
		}
		if (nextPulseDone==false)
		{
			Next_Pwd(pFilePwdList,&IndexFilePwdList,&IndexFilePwd,&IndexPwd);
			//pFilePwdList		= mProject.pFilePwdListSt;
		}
		//ul_PulseCount		= mProject.workSheetConfiguration.PulseQtyCriteria;
		do 
		{
			d_TimeEnd_ms		+= mProject.workSheetConfiguration.IntervalTimeCriteria;
		} while (d_TimeEnd_ms<=d_Time_ms);
	}
	while(IndexFilePwdList < mProject.FilePwdList_Count);
	
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
	
	mProject.pProjectFile	= new_SpreadFileList(&SpreadFileList);
	_SaveInfo(&mProject.pProjectFile->stPtdIni,&stPjtPtd);
	_SaveInfo(&mProject.pProjectFile->stPtdEnd,&stPjtPtdEnd);
	getPjtInfo			= true;
	mProject.pProjectFile->ul_PulseCount = ul_QtyPjt;
}

unsigned_long Pulse_project::Add_Pulses_To_Index(__PwdIndex* pPwdIndex,
												 __File_Pwd_List_St* pFilePwdList,
												 unsigned_long* IndFilePwdLst,
												 unsigned_long* IndFilePwd,
												 unsigned_long* IndPulse,
												 double* p_dTime_ms,
												 double d_TimeEnd_ms)
{
	__File_Pwd_St*	pFilePwdSt	= pFilePwdList->p_st_FilePwdArray+*IndFilePwd;
	__Pwd_NF_St*	pPwdNF = pFilePwdSt->p_st_Pwd_NewFields;
	unsigned_long	ul_Pulses	= pFilePwdSt->l_Pulse_Count - *IndPulse;
	unsigned_long	ul_Qty;
	unsigned_long	IndexIni = *IndPulse;
	unsigned_long	IndexEnd = pFilePwdSt->l_Pulse_Count-1;
	unsigned_long	Index;
// 	double			dInterval;
	
	*p_dTime_ms		= pPwdNF[IndexEnd].d_Rel_Toa_ms;
	if ( *p_dTime_ms> d_TimeEnd_ms)
	{
		do 
		{
			Index	= (IndexEnd + IndexIni) >> 1;
			*p_dTime_ms		= pPwdNF[Index].d_Rel_Toa_ms;
			if (*p_dTime_ms > d_TimeEnd_ms)
			{
				if (IndexEnd!=Index)
				{
					IndexEnd	= Index;
				} 
				else
				{
					break;
				}
			} 
			else
			{
				if (IndexIni!=Index)
				{
					IndexIni	= Index;
				} 
				else
				{
					break;
				}
			}
		} while (IndexIni!=IndexEnd);
		if (*IndPulse==Index)
		{
			(*IndPulse)		++;
			*p_dTime_ms		= pPwdNF[*IndPulse].d_Rel_Toa_ms;
		} 
		else
		{
			*IndPulse		= Index;		
		}
		set_pwd_index(pPwdIndex,*IndFilePwdLst,*IndFilePwd,*IndPulse);
		ul_Qty			= ul_Pulses - (pFilePwdSt->l_Pulse_Count - *IndPulse) + 1;
	} 
	else
	{
		(*IndPulse)		= IndexEnd;
		set_pwd_index(pPwdIndex,*IndFilePwdLst,*IndFilePwd,*IndPulse);
		ul_Qty			= ul_Pulses;
	}
	return ul_Qty;
}

double Pulse_project::TimeNextPulse(__File_Pwd_List_St* pFilePwdList,
									  unsigned_long IndFilePwdLst,
									  unsigned_long IndFilePwd,
									  unsigned_long IndPulse)
{
	return pFilePwdList->p_st_FilePwdArray[IndFilePwd].p_st_Pwd_NewFields[IndPulse].d_Rel_Toa_ms;
}
