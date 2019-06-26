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

//#include <.h>
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

void pulse_import::import_file(Pulse_project* pPulseProject)
{
	__ProjectSt*			pProject	= &pPulseProject->mProject;
	char* 					p_ch_Pulse_Path	= pProject->missionPath->FileName;
	__Ptd_St*				ptdStPrev = NULL;
	__Pwd_NF_St*			pPwdNfSt;
	__File_Pwd_St*			filePwdStructPrev = NULL;
	__File_Pwd_St*			pFilePwdSt;
	__File_List_St*			fileListStruct;
	__File_Pwd_List_St		filePwdLstStruct[2];
	__File_Pwd_List_St*		pFilePwdLstStNew = NULL;
	__File_Pwd_List_St*		pFilePwdLstStPrev= NULL;
	short					IndexNew=1;
	long					IndexLast;
	list<cFileName*>		cListFileName;
	cFilesList				cFileLst;
	file_list				c_file_list;
	long					lFileCount	= 0;
	unsigned long			i;
	double					dPostDToa_us;

	//_CrtDumpMemoryLeaks();
	c_file_list.get_file_list(p_ch_Pulse_Path, &cListFileName);
	pPulseProject->Set_FileCount(cListFileName.size());
	c_file_list.get_file_list(&cListFileName,
							  &cFileLst,
							  p_ch_Pulse_Path,
							  pProject->FilesPerWorkSpace,
							  &(pProject->FilePwdList_Count));

	fileListStruct			= cFileLst.mFileListHead;
	_InitFilePwdLstSt(filePwdLstStruct);
	_InitFilePwdLstSt(filePwdLstStruct+1);
	pProject->ul_PulseCount	= 0;
	cleanStartTime();
	m_ul_Correl_Index	= 1;
	m_ul_Relati_Index	= 1;
	m_p_d_pre__d_Toa_us	= NULL;
	m_p_dbg_d_pre__d_Toa_us	= NULL;
	m_ErrorPtrLst.clear();
	pPulseProject->ClearErrorFileCount();

	for (i=0; i<pProject->FilePwdList_Count; i++)
	{
		pPulseProject->Set_Index_FilePwdList(i);
		pPulseProject->Set_Index_FilePulse(lFileCount);
		
		//_CrtDumpMemoryLeaks();
		IndexNew	= 1 - IndexNew;
		pFilePwdLstStNew	= filePwdLstStruct+IndexNew;
		import_file_list(fileListStruct, 
						 ptdStPrev,
						 filePwdStructPrev,
						 pFilePwdLstStNew,
						 &lFileCount,
						 &pProject->ul_PulseCount,
						 &dPostDToa_us);
		//_CrtDumpMemoryLeaks();
		pFilePwdSt			= pFilePwdLstStNew->p_st_FilePwdArray+(pFilePwdLstStNew->us_ListCount-1);
		ptdStPrev			= &(pFilePwdSt->st_Last_Ptd);
		if (pFilePwdLstStPrev != NULL)
		{
			IndexLast		= pFilePwdLstStPrev->us_ListCount-1;
			pFilePwdSt		= pFilePwdLstStPrev->p_st_FilePwdArray+IndexLast;
			IndexLast		= pFilePwdSt->l_Pulse_Count-1;
			pPwdNfSt		= pFilePwdSt->p_st_Pwd_NewFields+IndexLast;
			//pPwdNfSt->d_pre__d_Toa_us	= dPostDToa_us;
			pPulseProject->set_filePwdListSt(pFilePwdLstStPrev,i-1);
			pPulseProject->save_filePwdListSt();
		}
		// Save Error Report
		pPulseProject->Set_ErrorPtrLstNew(m_ErrorPtrLstPt,m_ErrPrtCnt);
		pPulseProject->SaveErrPntLst();
		pPulseProject->UnSetErrPntLst();
		m_PoolPwdErr->Free(m_ErrorPtrLstPt);

		pFilePwdLstStPrev	= pFilePwdLstStNew;
		filePwdStructPrev	= pFilePwdLstStPrev->p_st_FilePwdArray+(pFilePwdLstStPrev->us_ListCount-1);
		fileListStruct		++;
	} // for (i=0; i<pProject->FilePwdList_Count; i++)
	pPulseProject->Set_Index_FilePwdList(i);
	pPulseProject->Set_Index_FilePulse(lFileCount);
	if (pFilePwdLstStPrev != NULL)
	{
		IndexLast		= pFilePwdLstStPrev->us_ListCount-1;
		pFilePwdSt		= pFilePwdLstStPrev->p_st_FilePwdArray+IndexLast;
		IndexLast		= pFilePwdSt->l_Pulse_Count-1;
		pPwdNfSt		= pFilePwdSt->p_st_Pwd_NewFields+IndexLast;
		pPwdNfSt->d_pre__d_Toa_us	= dPostDToa_us;
		pPulseProject->set_filePwdListSt(pFilePwdLstStPrev,i-1);
		pPulseProject->save_filePwdListSt();
		fileListStruct		++;
	}
	//m_Enlased_Tick		= ::GetTickCount() - m_Tick_Ini;
	pPulseProject->mProject.pFilePwdListSt	= NULL;
	_FieldDestroy(filePwdLstStruct);
	_FieldDestroy(filePwdLstStruct+1);
	c_file_list.DestroyList(&cListFileName);
	//pPulseProject->Set_ImportFileDone();

	/*pPulseProject->Set_ErrorPtrLstNew(m_ErrorPtrLstPt,m_ErrPrtCnt);
	pPulseProject->SaveErrPntLst();
	pPulseProject->UnSetErrPntLst();*/
	
	pPulseProject->Set_ImportFileDone();
	m_ErrorPtrLst.clear();
	//_CrtDumpMemoryLeaks();
}

void pulse_import::assign_pre_dtoa(__Pwd_NF_St* pPwdNfStSrc,__Pwd_NF_St* pPwdNfDst)
{
	pPwdNfDst->d_pre__d_Toa_us	= pPwdNfStSrc->d_post_d_Toa_us;
}

void pulse_import::assign_pre_dtoa(__File_Pwd_List_St* p_FilePwdLstStPrev, 
									__File_Pwd_List_St* p_FilePwdLstStNow)
{
	__Pwd_NF_St*	pSrc;
	__Pwd_NF_St*	pDst;
	long			i;
	__File_Pwd_St*	pFilePwdStPrev;

	pSrc			= p_FilePwdLstStNow->p_st_FilePwdArray->p_st_Pwd_NewFields;
	i				= p_FilePwdLstStPrev->us_ListCount-1;
	pFilePwdStPrev	= p_FilePwdLstStPrev->p_st_FilePwdArray + i;
	i				= pFilePwdStPrev->l_Pulse_Count - 1;
	pDst			= pFilePwdStPrev->p_st_Pwd_NewFields + i;
	assign_pre_dtoa(pSrc,pDst);
}

__File_Pwd_List_St	* pulse_import::import_file_list(__file_list_st *fileListStruct, 
													 __Ptd_St * p_LastPreviousTime,
													 long* pFileCounter)
{
	__File_Pwd_List_St		*filePwdLstStruct;
	struct __file_pwd_st	*filePwdStruct;
	struct __file_pwd_st	*filePwdStructPrev = NULL;
	__Ptd_St				*ptdStPrev = p_LastPreviousTime;
	unsigned long			i;
	double					d_FileTime_s;

	//filePwdLstStruct		= new __File_Pwd_List_St;
	filePwdLstStruct		= m_PoolFilePwdLst->Alloc();

	filePwdLstStruct->us_ListCount 		= fileListStruct->us_ListCount;;
	assign_path_name(filePwdLstStruct,fileListStruct);
	//filePwdLstStruct->p_st_FilePwdArray = new struct __file_pwd_st[filePwdLstStruct->us_ListCount] ;
	filePwdLstStruct->p_st_FilePwdArray = m_PoolFilePwd->Alloc(filePwdLstStruct->us_ListCount) ;
	filePwdLstStruct->p_st_FileList		= fileListStruct;

	filePwdStruct						= filePwdLstStruct->p_st_FilePwdArray;
	filePwdLstStruct->ul_PulseCount		= 0;
	for(i = 0; i < fileListStruct->us_ListCount; i++)
	{
		filePwdStruct->s_WrapAround_Counter	= 0;
		filePwdStruct->uc_OverFlow_Flag		= 0;
		//filePwdStruct->p_ch_FileName		= fileListStruct->p_ch_FileList[i];
		//filePwdStruct->us_FileName_length	= fileListStruct->p_us_NamesLenList[i];
		assign_file_name					(filePwdStruct,fileListStruct,i);
		read_file							(filePwdStruct, filePwdLstStruct->p_PathName);
		d_FileTime_s						= get_time_file_s(filePwdStruct);
		if (!i)
		{
			m_d_Start_Time_s				= d_FileTime_s;
		}
		get_first_ptd						(filePwdStruct, d_FileTime_s, ptdStPrev);
		ptdStPrev							= &(filePwdStruct->st_Last_Ptd);
		get_pwd_nf							(filePwdStruct,filePwdStructPrev,i);
		get_last_ptd						(filePwdStruct);
		filePwdLstStruct->ul_PulseCount		+= filePwdStruct->l_Pulse_Count;
		filePwdStruct->l_FileIndex			= (*pFileCounter)++;
		filePwdStructPrev					= filePwdStruct;
		filePwdStruct						++;
	}
	filePwdLstStruct->d_TotalTime_ms		= filePwdStructPrev->st_Last_Ptd.d_Time_ms;
	destroy_File_Pwd_St(m_File_Pwd_List_St_Recent);
	m_File_Pwd_List_St_Recent = filePwdLstStruct;
	return filePwdLstStruct;
}

void pulse_import::assign_path_name(__File_Pwd_List_St* p_FilePwdLstSt,
									 __File_List_St * p_File_List_St)
{
	p_FilePwdLstSt->s_PathNameLenght	= strlen(p_File_List_St->p_ch_Path);
	p_FilePwdLstSt->p_PathName			= new char[p_FilePwdLstSt->s_PathNameLenght+1];
	strcpy(p_FilePwdLstSt->p_PathName, p_File_List_St->p_ch_Path);
}

void pulse_import::assign_file_name(__file_pwd_st * p_st_FilePwd,
									 __File_List_St * p_File_List_St,
									 long Index)
{
	p_st_FilePwd->us_FileName_length	= p_File_List_St->p_us_NamesLenList[Index];
	p_st_FilePwd->p_ch_FileName	= new char[p_File_List_St->p_us_NamesLenList[Index]+1];
	strcpy(p_st_FilePwd->p_ch_FileName, p_File_List_St->p_ch_FileList[Index]);
}

void pulse_import::_InitFilePwdLstSt(__File_Pwd_List_St* pFilePwdLst)
{
	pFilePwdLst->p_PathName			= NULL;
	pFilePwdLst->p_st_FileList		= NULL;
	pFilePwdLst->p_st_FilePwdArray	= NULL;
}

void pulse_import::setStartTime(double NewStartTime)
{
	if (m_d_StartTimeEmpty==true)
	{
		m_d_StartTimeEmpty	= false;
		m_d_Start_Time_s	= NewStartTime;
	}
}