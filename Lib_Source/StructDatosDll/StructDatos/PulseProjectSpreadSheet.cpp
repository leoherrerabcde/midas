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


// long Pulse_project::GetPwdCount(long IndexSpread , 
// 					long IndexSheet ,
// 					long Index)
// {
// 	__WorkSheetBounds	*pWorkSheetBounds;
// 
// 	return pWorkSheetBounds->ul_PulseCount;
// }

void Pulse_project::GetSheetInfo(long IndexSpread , 
					 long IndexSheet ,
					 long * PulseQty ,
					 double* TimeIni ,
					 double* TimeEnd)
{
	__WorkSheetBounds*	pWorkSheetBouns;

	pWorkSheetBouns	= mProject.pProjectFile->pSpreadFileArray[IndexSpread].pWorkSheetArray+IndexSheet;
	
	*PulseQty		= pWorkSheetBouns->ul_PulseCount;
	*TimeEnd		= pWorkSheetBouns->stPtdEnd.d_Time_ms;
	*TimeIni		= pWorkSheetBouns->stPtdIni.d_Time_ms;
}

void Pulse_project::setColumnFormat(long* ColumnOrder, long* ColumnEnable)
{
	memcpy(mProject.workSheetConfiguration.ColumnOrder,ColumnOrder,PWD_FIELD_COUNT*sizeof(long));
	memcpy(mProject.workSheetConfiguration.ColumnEnable,ColumnEnable,PWD_FIELD_COUNT*sizeof(long));
}


void Pulse_project::GetSpreadinfo(long IndexSpread , 
					  long * PulseQty ,
					  double* TimeIni ,
					  double* TimeEnd)
{
	__SpreadFile*	pSpreadFile;

	pSpreadFile		= mProject.pProjectFile->pSpreadFileArray+IndexSpread;
	*PulseQty		= pSpreadFile->us_PulseCount;
	*TimeEnd		= pSpreadFile->stPtdEnd.d_Time_ms;
	*TimeIni		= pSpreadFile->stPtdIni.d_Time_ms;
}

void Pulse_project::CreateSheet (long IndexSpread,long IndexSheet)
{

}

void Pulse_project::SaveSpreadSheet (long IndexSpread)
{

}

