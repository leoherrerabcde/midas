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

void pulse_import::import_file_v2(Pulse_project* pPulseProject)
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

	pPulseProject->Set_MaxPulses(pPulseProject->Get_FileMaxPulses(fileListStruct,
								pProject->FilePwdList_Count));

	_InitFilePwdLstSt(filePwdLstStruct);
	_InitFilePwdLstSt(filePwdLstStruct+1);
	pProject->ul_PulseCount	= 0;
	cleanStartTime();
	m_ul_Correl_Index	= 1;
	m_ul_Relati_Index	= 1;
	m_p_d_pre__d_Toa_us	= NULL;
	m_p_dbg_d_pre__d_Toa_us	= NULL;
	//m_Tick_Ini			= ::GetTickCount();
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
	pPulseProject->Set_ImportFileDone();
	//_CrtDumpMemoryLeaks();
}

