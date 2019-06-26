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
#include "ErrorHandler.h"

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

extern CErrorHandler		mErrorHandler;

void pulse_import::import_file_optimize(Pulse_project* pPulseProject)
{
	__Pwd_St				lvLastPwd;
	__Ptd_St*				ptdStPrev			= NULL;
	__Pwd_NF_St				lvLastPwdNfSt;
	__Pwd_NF_St*			pPwdNfSt;
	__ProjectSt*			pProject			= &pPulseProject->mProject;
	__File_Pwd_St*			filePwdStructPrev	= NULL;
	__File_Pwd_St*			pFilePwdSt;
	__File_Pwd_St			lvLastFilePwdSt;
	__File_List_St*			fileListStruct;
	//__File_Pwd_List_St		filePwdLstStruct[2];
	__File_Pwd_List_St*		pFilePwdLstSt		= NULL;
	__File_Pwd_List_St*		pFilePwdLstStNew	= NULL;
	__File_Pwd_List_St*		pFilePwdLstStPrev	= NULL;
	__File_Pwd_List_St*		pFilePwdLstStDestroy= NULL;
	char* 					p_ch_Pulse_Path		= pProject->missionPath->FileName;

	//short					IndexNew			= 1;
	long					IndexLast;
	long					lFileCount			= 0;
	unsigned long			i;
	double					dPostDToa_us;
	file_list				c_file_list;
	cFilesList				cFileLst;
	list<cFileName*>		cListFileName;
	
	c_file_list.get_file_list(p_ch_Pulse_Path, &cListFileName);
	pPulseProject->Set_FileCount(cListFileName.size());
	c_file_list.get_file_list(&cListFileName,
								&cFileLst,
								p_ch_Pulse_Path,
								pProject->FilesPerWorkSpace,
								&(pProject->FilePwdList_Count));
	/*c_file_list.get_file_list_op(&cListFileName,
								&cFileLst,
								p_ch_Pulse_Path,
								pProject->FilesPerWorkSpace,
								&(pProject->FilePwdList_Count));*/

	pFilePwdLstSt	= SetFileListStArray(pPulseProject,&cFileLst,p_ch_Pulse_Path);

	fileListStruct			= cFileLst.mFileListHead;

	cleanStartTime();
	m_ul_Correl_Index	= 1;
	m_ul_Relati_Index	= 1;
	m_p_d_pre__d_Toa_us	= NULL;
	m_p_dbg_d_pre__d_Toa_us	= NULL;
	m_ErrorPtrLst.clear();
	pFilePwdLstStNew	= pFilePwdLstSt;
	InitFilePwdSt(lvLastFilePwdSt,lvLastPwd,lvLastPwdNfSt);

	for (i=0; i<pProject->FilePwdList_Count; i++)
	{
		pPulseProject->Set_Index_FilePwdList(i);
		pPulseProject->Set_Index_FilePulse(lFileCount);
		DestroyFilePwdLstSt_op(pFilePwdLstStDestroy);
		
		import_file_list_op(fileListStruct, 
						ptdStPrev,
						filePwdStructPrev,
						pFilePwdLstStNew,
						&lFileCount,
						&pProject->ul_PulseCount,
						&dPostDToa_us);
		pFilePwdSt		= pFilePwdLstStNew->p_st_FilePwdArray+
						(pFilePwdLstStNew->us_ListCount-1);
		ptdStPrev		= &(pFilePwdSt->st_Last_Ptd);
		if (pFilePwdLstStPrev != NULL)
		{
			IndexLast		= pFilePwdLstStPrev->us_ListCount-1;
			pFilePwdSt		= pFilePwdLstStPrev->p_st_FilePwdArray+IndexLast;
			IndexLast		= pFilePwdSt->l_Pulse_Count-1;
			pPwdNfSt		= pFilePwdSt->p_st_Pwd_NewFields+IndexLast;
			pPulseProject->set_filePwdListSt(pFilePwdLstStPrev,i-1);
			pPulseProject->save_filePwdListSt();
		}
		pFilePwdLstStPrev	= pFilePwdLstStNew;
		filePwdStructPrev	= pFilePwdLstStPrev->p_st_FilePwdArray+(pFilePwdLstStPrev->us_ListCount-1);
		fileListStruct		++;
		pFilePwdLstStNew	++;
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
	//_FieldDestroy(filePwdLstStruct);
	//_FieldDestroy(filePwdLstStruct+1);
	c_file_list.DestroyList(&cListFileName);
	pPulseProject->Set_ImportFileDone();
	//pPulseProject->Set_ErrorPtrLst(&m_ErrorPtrLst);
	pPulseProject->Set_ErrorPtrLst(NULL);
	pPulseProject->SaveErrPntLst();
	pPulseProject->DestroyErrPntLst();
	pPulseProject->Set_ImportFileDone();
	m_ErrorPtrLst.clear();
}

__File_Pwd_List_St* pulse_import::SetFileListStArray(Pulse_project* pPulseProject,
									  cFilesList* p_cFileList,
									  char *pPath)
{

	__ProjectSt*			pProject			= &pPulseProject->mProject;
	unsigned long			i;
	__File_Pwd_List_St*		filePwdLstStruct;
	__File_Pwd_List_St*		pFilePwdLstStNew	= NULL;
	__File_Pwd_List_St*		pFilePwdLstStPrev	= NULL;
	__File_List_St*			pFileLstSt;

	m_FileIndex				= 0;
	pProject->ul_PulseCount	= 0;

	if (m_PoolFilePwdLst->CheckSpaceAvailable(pProject->FilePwdList_Count)==false)
	{
		mErrorHandler.AddError(1,__FILE__,"pulse_import::SetFileListStArray()","Not Enough Space for filePwdLstStruct");
		return NULL;
	}
	filePwdLstStruct		= m_PoolFilePwdLst->Alloc(pProject->FilePwdList_Count);
	pFilePwdLstStNew		= filePwdLstStruct;
	pFileLstSt				= p_cFileList->mFileListHead;
	pProject->pFilePwdListSt= filePwdLstStruct;

	for (i=0; i<pProject->FilePwdList_Count; i++)
	{
		SetFilePwdLstSt(pFilePwdLstStNew,pFileLstSt);
		pProject->ul_PulseCount	+= pFilePwdLstStNew->ul_PulseCount;
		pFileLstSt			++;
		pFilePwdLstStNew	++;
	}
	return	filePwdLstStruct;
}

void pulse_import::SetFilePwdLstSt(__File_Pwd_List_St* pFilePwdLstSt,__File_List_St* pFileLstSt)
{
	char**				pFileNameList;
	__File_Pwd_St*		pFilePwdSt;
	unsigned_long		i;

	pFilePwdLstSt->us_ListCount			= pFileLstSt->us_ListCount;
	pFilePwdLstSt->p_st_FileList		= pFileLstSt;
	pFilePwdSt							= m_PoolFilePwd->Alloc(pFilePwdLstSt->us_ListCount);
	pFilePwdLstSt->p_st_FilePwdArray	= pFilePwdSt;
	pFileNameList						= pFileLstSt->p_ch_FileList;
	pFilePwdLstSt->ul_PulseCount		= 0;
	pFilePwdLstSt->p_PathName			= pFileLstSt->p_ch_Path;
	pFilePwdLstSt->s_PathNameLenght		= strlen(pFileLstSt->p_ch_Path);

	for (i=0;i<pFilePwdLstSt->us_ListCount;i++)
	{
		SetFilePwd(pFilePwdSt,*pFileNameList,pFileLstSt->p_ch_Path);
		pFilePwdSt->us_FileName_length	= pFileLstSt->p_us_NamesLenList[i];
		pFilePwdLstSt->ul_PulseCount	+= pFilePwdSt->l_Pulse_Count;
		pFilePwdSt		++;
		pFileNameList	++;
	}
}

void pulse_import::SetFilePwd(__File_Pwd_St* pFilePwdSt,char *pFileName,char* p_ch_Path)
{
	pFilePwdSt->l_Pulse_Count		= Get_FilePulses(p_ch_Path,pFileName);
	pFilePwdSt->l_FileIndex			= m_FileIndex++;
	pFilePwdSt->p_ch_FileName		= pFileName;
	pFilePwdSt->s_WrapAround_Counter= 0;
	pFilePwdSt->uc_OverFlow_Flag	= 0;
	pFilePwdSt->p_st_Pwd			= NULL;
	pFilePwdSt->p_st_Pwd_NewFields	= NULL;
}


unsigned_long pulse_import::Get_FilePulses(char* pPath,char* pFileName)
{
	FILE			*pFile;
	long			l_FileSize;
	char			FileName[300];
	
	strcpy(FileName,pPath);
	strcat(FileName,"\\");
	strcat(FileName,pFileName);

	pFile               = fopen(FileName,"rb");
	if (pFile == NULL)
	{
		return 0;
	}
	fseek( pFile, 0, SEEK_END );
	l_FileSize	=  ftell( pFile );
	fclose(pFile);
	return l_FileSize / sizeof(__Pwd_St);
}

void pulse_import::InitFilePwdSt(__File_Pwd_St &pFilePwdSt,
								 __Pwd_St &pPwdSt,
								 __Pwd_NF_St &pPwdNFSt)
{
	pFilePwdSt.l_Pulse_Count		= 1;
	pFilePwdSt.p_st_Pwd				= &pPwdSt;
	pFilePwdSt.p_st_Pwd_NewFields	= &pPwdNFSt;
	//pFilePwdSt.st_First_Ptd			= pPtdStIni;
	//pFilePwdSt.st_Last_Ptd			= pPtdStEnd;
}

