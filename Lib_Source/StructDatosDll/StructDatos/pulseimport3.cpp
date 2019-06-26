/*
 * pulseimport.cpp
 *
 *  Created on: Sep 16, 2012
 *      Author: lherrera
 */

#include "pulse_conv_struct_define.h"
#include "pulseimport.h"
#include "filelist.h"
#include "cFileName.h"
#include "cFilesList.h"

#include <stdio.h>
#include <stdlib.h>
#include <time.h>
#include <dirent.h>
#include <string.h>
#include <iostream>
#include <list>
using namespace std;

#ifdef _MSC_VER
#include <crtdbg.h>
#endif



void pulse_import::import_file_list(__file_list_st *fileListStruct,
									__Ptd_St * p_LastPreviousTime,
									__File_Pwd_St* filePwdStructPrev,
									__File_Pwd_List_St	*filePwdLstStruct,
									long* pFileCounter,
									unsigned_long* pCountPulse,
									double* pLastPostDToaus)
{
	__File_Pwd_St*			filePwdStruct;
	//struct __file_pwd_st	*filePwdStructPrev = NULL;
	__Ptd_St				*ptdStPrev = p_LastPreviousTime;
	unsigned long			i;
	double					d_FileTime_s;
	__Pwd_NF_St*			pPwdNfSt;
	
	_DestroyArraySt(filePwdLstStruct->p_st_FilePwdArray,filePwdLstStruct->us_ListCount);

	filePwdLstStruct->us_ListCount 		= fileListStruct->us_ListCount;;
	assign_path_name(filePwdLstStruct,fileListStruct);

	//filePwdLstStruct->p_st_FilePwdArray = new struct __file_pwd_st[filePwdLstStruct->us_ListCount] ;
	m_PoolFilePwd->VerifyAllocated(filePwdLstStruct->p_st_FilePwdArray);
	filePwdLstStruct->p_st_FilePwdArray = m_PoolFilePwd->Alloc(filePwdLstStruct->us_ListCount);
	
	filePwdLstStruct->p_st_FileList		= fileListStruct;
	
	m_ErrPrtCnt							= 0;
	m_ErrPrtMaxCnt						= m_PoolPwdErr->GetSizePoolTable()/2;
	m_ErrorPtrLstPt						= m_PoolPwdErr->Alloc(m_ErrPrtMaxCnt);
	if (m_ErrorPtrLstPt==NULL)
	{
		m_ErrorPtrLstPt=m_ErrorPtrLstPt;
	}

	filePwdStruct						= filePwdLstStruct->p_st_FilePwdArray;
	filePwdLstStruct->ul_PulseCount		= 0;
	for(i = 0; i < fileListStruct->us_ListCount; i++)
	{
		filePwdStruct->s_WrapAround_Counter	= 0;
		filePwdStruct->uc_OverFlow_Flag		= 0;
		assign_file_name					(filePwdStruct,fileListStruct,i);
		read_file							(filePwdStruct, filePwdLstStruct->p_PathName);
		//_CrtDumpMemoryLeaks();
		d_FileTime_s						= get_time_file_s(filePwdStruct);
		setStartTime(d_FileTime_s);
		get_first_ptd						(filePwdStruct, d_FileTime_s, ptdStPrev);
		ptdStPrev							= &(filePwdStruct->st_Last_Ptd);
		get_pwd_nf							(filePwdStruct,filePwdStructPrev,*pFileCounter);
		//_CrtDumpMemoryLeaks();
		get_last_ptd						(filePwdStruct);
		filePwdLstStruct->ul_PulseCount		+= filePwdStruct->l_Pulse_Count;
		filePwdStruct->l_FileIndex			= (*pFileCounter)++;
		filePwdStructPrev					= filePwdStruct;
		filePwdStruct						++;
	}
	filePwdLstStruct->d_TotalTime_ms		= filePwdStructPrev->st_Last_Ptd.d_Time_ms;
	i										= filePwdStructPrev->l_Pulse_Count - 1;
	pPwdNfSt								= filePwdStructPrev->p_st_Pwd_NewFields+i;
	*pLastPostDToaus						= pPwdNfSt->d_post_d_Toa_us;
	*pCountPulse							+= filePwdStruct->l_Pulse_Count;
	//m_PoolPwdErr->Free(m_ErrorPtrLstPt);
}

