/*
 * Pulseproject.cpp
 *
 *  Created on: Oct 5, 2012
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

Pulse_project::Pulse_project() {
	// TODO Auto-generated constructor stub
	long	i=0;
	mProject.missionPath	= NULL;
	mProject.workSpacePath	= NULL;
	mProject.outputPath		= NULL;

	mProject.IndexFilePwdLstSt	= -1;
	mProject.IndexFilePwdSt		= -1;

	mProject.pFilePwdListSt	= NULL;
	mProject.pFilePwdSt		= NULL;
	mProject.pProjectFile	= NULL;
	mProject.pFileListSt	= NULL;

	mProject.TempIndex.IndexSheet	= -1;
	mProject.TempIndex.IndexSpread	= -1;
	
	mProject.mErrPntList.Count			= 0;
	mProject.mErrPntList.PointerArray	= NULL;

	for (i=0;i<PWD_FIELD_COUNT;i++)
	{
		mProject.workSheetConfiguration.ColumnOrder[i]=i;
		mProject.workSheetConfiguration.ColumnEnable[i]=1;
	}

	m_PoolFilePwdLst		= NULL;
	m_PoolFilePwd			= NULL;
	m_PoolPwdNF				= NULL;
	m_PoolPwd				= NULL;

	m_ImportFileDone		= 0;
}

Pulse_project::~Pulse_project() {
	// TODO Auto-generated destructor stub

	DestroyFileNameSt(mProject.missionPath);
	DestroyFileNameSt(mProject.workSpacePath);
	DestroyFileNameSt(mProject.outputPath);
	_Destroy(mProject.pFileListSt);
	_Destroy(mProject.pFilePwdListSt);
	_Destroy(mProject.pFilePwdSt);
	_Destroy(mProject.pProjectFile);
}

void Pulse_project::DestroyFileNameSt(__FileNameSt* pFileNameSt)
{
	if (pFileNameSt != NULL)
	{
		delete []pFileNameSt->FileName;
		delete pFileNameSt;
	}
}

bool	Pulse_project::write_in_file(void * pSource,size_t size, size_t count, FILE* pFile)
{
	size_t		byteWritten;

	byteWritten	= fwrite(pSource,size,count,pFile);

	if (byteWritten != count)
	{
		return false;
	}
	return true;
}

bool	Pulse_project::read_in_file(void * pDst,size_t size, size_t count, FILE* pFile)
{
	size_t		byteRead;
	
	byteRead	= fread(pDst,size,count,pFile);
	
	if (byteRead != count)
	{
		return false;
	}
	return true;
}

void	Pulse_project::get_fileNameFilePwdLstSt(char* pFileName)
{
	//char	Tmp[MAX_LEN_FILENAME];
	long	len_filename;

	if (mProject.workSpacePath!=NULL)
	{
	len_filename = strlen(mProject.workSpacePath->FileName);
	if (len_filename>260-13)
	{
		len_filename++;
	}
	sprintf(pFileName,"%s\\tmp%05d.fpl",
			mProject.workSpacePath->FileName,
			mProject.IndexFilePwdLstSt);
	}
}

void	Pulse_project::get_fileNameFilePwdLstSt(char* pFileName, long index)
{
	mProject.IndexFilePwdLstSt = index;
	get_fileNameFilePwdLstSt(pFileName);
}

bool	Pulse_project::save_filePwdListSt(void)
{
	FILE*				pFile;
	char				FileName[260];
	__File_Pwd_List_St*	pFilePwdLstSt;
	__File_Pwd_St*		pFilePwdSt;
	__Pwd_St*			pPwdSt;
	__Pwd_NF_St*		pPwdNFSt;
	char*				pPulseFileName;
	//char*				pPath;
	long				Index;

	pFilePwdLstSt		= mProject.pFilePwdListSt;
	pFilePwdSt			= pFilePwdLstSt->p_st_FilePwdArray;

	get_fileNameFilePwdLstSt(FileName);
	pFile		= fopen(FileName,"wb");
	if (pFile==NULL)
	{
		return false;
	}
	if(write_in_file(pFilePwdLstSt,sizeof(__File_Pwd_List_St),1,pFile)==false)
	{
		return false;
	}
	if(write_in_file(pFilePwdSt,sizeof(__File_Pwd_St),pFilePwdLstSt->us_ListCount,pFile)==false)
	{
		return false;
	}
	for(Index=0;Index<pFilePwdLstSt->us_ListCount;Index++)
	{
		pPwdSt			= pFilePwdSt->p_st_Pwd;
		pPwdNFSt		= pFilePwdSt->p_st_Pwd_NewFields;
		pPulseFileName	= pFilePwdSt->p_ch_FileName;
		if(write_in_file(pPwdSt,sizeof(__Pwd_St),pFilePwdSt->l_Pulse_Count,pFile)==false)
		{
			return false;
		}
		if(write_in_file(pPwdNFSt,sizeof(__Pwd_NF_St),pFilePwdSt->l_Pulse_Count,pFile)==false)
		{
			return false;
		}
		if(write_in_file(pPulseFileName,sizeof(char),pFilePwdSt->us_FileName_length+1,pFile)==false)
		{
			return false;
		}
		pFilePwdSt ++;
	}
	fclose(pFile);
	return true;
}

bool	Pulse_project::read_file(long Index)
{
	FILE*				pFile;
	char				FileName[260];
	__File_Pwd_List_St*	pFilePwdLstSt;
	__File_Pwd_St*		pFilePwdSt;
	__Pwd_St*			pPwdSt;
	__Pwd_NF_St*		pPwdNFSt;
	char*				pPulseFileName;
	long				i;
	
	get_fileNameFilePwdLstSt(FileName,Index);
	pFilePwdLstSt	= mProject.pFilePwdListSt;
	
	pFile		= fopen(FileName,"rb");
	if (pFile==NULL)
	{
		return false;
	}
	//_CrtDumpMemoryLeaks();
	for (i=0;i<pFilePwdLstSt->us_ListCount;i++)
	{
		_DestroyFieldFilePwdSt(pFilePwdLstSt->p_st_FilePwdArray+i);
	}
	delete [] pFilePwdLstSt->p_st_FilePwdArray;
	//_DestroyFieldFilePwdSt(pFilePwdLstSt->p_st_FilePwdArray);
	//_CrtDumpMemoryLeaks();
	if(read_in_file(pFilePwdLstSt,sizeof(__File_Pwd_List_St),1,pFile)==false)
	{
		return false;
	}
	pFilePwdLstSt->p_PathName			= NULL;
	pFilePwdLstSt->p_st_FileList		= NULL;
	pFilePwdSt							= new __File_Pwd_St[pFilePwdLstSt->us_ListCount];
	pFilePwdLstSt->p_st_FilePwdArray	= pFilePwdSt;

	if(read_in_file(pFilePwdSt,sizeof(__File_Pwd_St),pFilePwdLstSt->us_ListCount,pFile)==false)
	{
		return false;
	}
	

	for(Index=0;Index<pFilePwdLstSt->us_ListCount;Index++)
	{
		pPwdSt							= new __Pwd_St[pFilePwdSt->l_Pulse_Count];
		pPwdNFSt						= new __Pwd_NF_St[pFilePwdSt->l_Pulse_Count];
		pPulseFileName					= new char[pFilePwdSt->us_FileName_length+1];
		pFilePwdSt->p_st_Pwd			= pPwdSt;
		pFilePwdSt->p_st_Pwd_NewFields	= pPwdNFSt;
		pFilePwdSt->p_ch_FileName		= pPulseFileName;
		if(read_in_file(pPwdSt,sizeof(__Pwd_St),pFilePwdSt->l_Pulse_Count,pFile)==false)
		{
			return false;
		}
		if(read_in_file(pPwdNFSt,sizeof(__Pwd_NF_St),pFilePwdSt->l_Pulse_Count,pFile)==false)
		{
			return false;
		}
		if(read_in_file(pPulseFileName,sizeof(char),pFilePwdSt->us_FileName_length+1,pFile)==false)
		{
			return false;
		}
		pFilePwdSt ++;
	}
	fclose(pFile);
	pFilePwdLstSt->p_st_FileList	= NULL;
	return true;
}

bool	Pulse_project::read_filePwdListSt(long Index)
{
	FILE*				pFile;
	char				FileName[260];
	__File_Pwd_List_St*	pFilePwdLstSt;
	__File_Pwd_St*		pFilePwdSt;
	__Pwd_St*			pPwdSt;
	__Pwd_NF_St*		pPwdNFSt;
	char*				pPulseFileName;
	//char*				pPath;
	
	get_fileNameFilePwdLstSt(FileName,Index);
	_DestroyFilePwdList(mProject.pFilePwdListSt);
	mProject.pFilePwdListSt	= new __File_Pwd_List_St;
	pFilePwdLstSt	= mProject.pFilePwdListSt;
	pFilePwdLstSt->p_st_FilePwdArray	= 0;
	pFilePwdLstSt->ul_PulseCount		= 0;
	pFilePwdLstSt->us_ListCount		= 0;
	pFilePwdLstSt->p_PathName		= NULL;
	pFilePwdLstSt->p_st_FileList		= NULL;
	pFilePwdLstSt->p_st_FilePwdArray	= NULL;

	pFile		= fopen(FileName,"rb");
	if (pFile==NULL)
	{
		return false;
	}

	
	if(read_in_file(pFilePwdLstSt,sizeof(__File_Pwd_List_St),1,pFile)==false)
	{
		return false;
	}
	
	pFilePwdSt							= new __File_Pwd_St[pFilePwdLstSt->us_ListCount];
	pFilePwdLstSt->p_st_FilePwdArray	= pFilePwdSt;
	pFilePwdLstSt->p_PathName			= NULL;
	pFilePwdLstSt->p_st_FileList		= NULL;
	if(read_in_file(pFilePwdSt,sizeof(__File_Pwd_St),pFilePwdLstSt->us_ListCount,pFile)==false)
	{
		return false;
	}


	for(Index=0;Index<pFilePwdLstSt->us_ListCount;Index++)
	{
		pPwdSt							= new __Pwd_St[pFilePwdSt->l_Pulse_Count];
		pPwdNFSt						= new __Pwd_NF_St[pFilePwdSt->l_Pulse_Count];
		pPulseFileName					= new char[pFilePwdSt->us_FileName_length+1];
		pFilePwdSt->p_st_Pwd			= pPwdSt;
		pFilePwdSt->p_st_Pwd_NewFields	= pPwdNFSt;
		pFilePwdSt->p_ch_FileName		= pPulseFileName;
		if(read_in_file(pPwdSt,sizeof(__Pwd_St),pFilePwdSt->l_Pulse_Count,pFile)==false)
		{
			return false;
		}
		if(read_in_file(pPwdNFSt,sizeof(__Pwd_NF_St),pFilePwdSt->l_Pulse_Count,pFile)==false)
		{
			return false;
		}
		if(read_in_file(pPulseFileName,sizeof(char),pFilePwdSt->us_FileName_length+1,pFile)==false)
		{
			return false;
		}
		pFilePwdSt ++;
	}
	fclose(pFile);
	pFilePwdLstSt->p_st_FileList	= NULL;
	return true;
}

void Pulse_project::getMissionName(char* lvName)
{
	char MissionName[260];

	strcpy(MissionName,mProject.missionPath->FileName);
	MissionName[mProject.missionPath->LengthString-23]='\0';
	strcpy(lvName,MissionName);
}
void Pulse_project::setMissionPath(__FileNameSt* pPath)
{
	DestroyFileNameSt(mProject.missionPath);
	mProject.missionPath	= new __FileNameSt;
	cpy_file_name(mProject.missionPath,pPath);
}

void Pulse_project::setWrkSpcPath(__FileNameSt* pPath)
{
	DestroyFileNameSt(mProject.workSpacePath);
	mProject.workSpacePath	= new __FileNameSt;
	cpy_file_name(mProject.workSpacePath,pPath);
}

void Pulse_project::setOutputPath(__FileNameSt* pPath)
{
	DestroyFileNameSt(mProject.outputPath);
	mProject.outputPath	= new __FileNameSt;
	cpy_file_name(mProject.outputPath,pPath);
}

void Pulse_project::cpy_file_name(__FileNameSt* Dst, __FileNameSt* Src)
{
	Dst->LengthString	= Src->LengthString;
	Dst->FileName		= new char[Dst->LengthString+1];
	strcpy(Dst->FileName,Src->FileName);
}

void Pulse_project::setIntervalPerSheet	(double Interval)
{
	__workSheetConfig	*workSheetConfig;
	workSheetConfig		= &(mProject.workSheetConfiguration);

	workSheetConfig->PulseQtyCriteria		= 0;
	workSheetConfig->IntervalTimeCriteria	= Interval;
}

void Pulse_project::setSheetsPerXls		(long SheetCount)
{
	__workSheetConfig	*workSheetConfig;
	workSheetConfig		= &(mProject.workSheetConfiguration);

	workSheetConfig->workSheetsPerXlsCount	= SheetCount;
}

void Pulse_project::setPulsesPerSheet	(long PulseCount)
{
	__workSheetConfig	*workSheetConfig;
	workSheetConfig		= &(mProject.workSheetConfiguration);

	workSheetConfig->PulseQtyCriteria		= PulseCount;
	workSheetConfig->IntervalTimeCriteria	= 0;
}

void Pulse_project::create_workspace	(void)
{
	__workSheetConfig	*workSheetConfig;
	workSheetConfig		= &(mProject.workSheetConfiguration);

	if (workSheetConfig->IntervalTimeCriteria > 0)
	{
		create_workspace_byTime();
	} 
	else
	{
		create_workspace_byPulse();
	}
}

void Pulse_project::create_workspace_byTime	(void)
{

}

void Pulse_project::set_pwd_index(__PwdIndex* pPwdIndex,
								  unsigned_long IndFilePwdLst,
							 	  unsigned_long IndFilePwd,
								  unsigned_long IndPulse)
{
	pPwdIndex->ul_IndexWorkSpace	= IndFilePwdLst;
	pPwdIndex->us_IndexFilePwd		= IndFilePwd;
	pPwdIndex->us_IndexPulse		= IndPulse;

}

void Pulse_project::Init_PwdIndex	(__PwdIndex* pPwdIndex)
{
	pPwdIndex->us_IndexSheet		= 0;
	pPwdIndex->ul_IndexWorkSpace	= 0;
	pPwdIndex->us_IndexFilePwd		= 0;
	pPwdIndex->us_IndexPulse		= 0;
}

void Pulse_project::PwdIndex_Cpy(__PwdIndex* pPwdIndexDst,__PwdIndex* pPwdIndexSrc)
{
	memcpy(pPwdIndexDst,pPwdIndexSrc,sizeof(__PwdIndex));
}

void Pulse_project::Add_Index_Ini(list<__PwdIndex*>* pListIndex,
								  __PwdIndex* pPwdIndex)
{
	__PwdIndex*	pNewPwdIndex	= new __PwdIndex;
	PwdIndex_Cpy(pNewPwdIndex,pPwdIndex);
	pListIndex->push_back(pNewPwdIndex);
}

void Pulse_project::Add_Index_Ini(list<__PwdIndex*>* pListIndex,
								  unsigned_long IndFilePwdLst,
								  unsigned_long IndFilePwd,
								  unsigned_long IndPulse)
{
	__PwdIndex* pPwdIndex = new __PwdIndex;
	set_pwd_index(pPwdIndex,IndFilePwdLst,IndFilePwd,IndPulse);
	pListIndex->push_back(pPwdIndex);
}

void Pulse_project::Add_Index_Ini(list<__PwdIndex*>* pListIndex,
								  __PwdIndex* pPwdIndex,
								  unsigned_long IndFilePwdLst,
								  unsigned_long IndFilePwd,
								  unsigned_long IndPulse)
{
	//pPwdIndex	= new __PwdIndex;
	set_pwd_index(pPwdIndex,IndFilePwdLst,IndFilePwd,IndPulse);
	pListIndex->push_back(pPwdIndex);
}

void Pulse_project::Add_Index_End(list<__PwdIndex*>* pListIndex,
								  unsigned_long IndFilePwdLst,
								  unsigned_long IndFilePwd,
								  unsigned_long IndPulse)
{
	__PwdIndex* pPwdIndex	= new __PwdIndex;
	set_pwd_index(pPwdIndex,IndFilePwdLst,IndFilePwd,IndPulse);
	pListIndex->push_back(pPwdIndex);
}


unsigned_long Pulse_project::Add_Pulses_To_Index(__PwdIndex* pPwdIndex,
							   __File_Pwd_List_St* pFilePwdList,
							   unsigned_long* IndFilePwdLst,
							   unsigned_long* IndFilePwd,
							   unsigned_long* IndPulse,
							   unsigned_long* pPulses)
{
	__File_Pwd_St*	pFilePwdSt	= pFilePwdList->p_st_FilePwdArray+*IndFilePwd;
	unsigned_long	ul_Pulses	= pFilePwdSt->l_Pulse_Count - *IndPulse;
	unsigned_long	ul_Qty;

	if (!*pPulses)
	{
		*pPulses	= ul_Pulses;
	}
	if (ul_Pulses > *pPulses)
	{
		(*IndPulse)		+= (*pPulses-1);
		set_pwd_index(pPwdIndex,*IndFilePwdLst,*IndFilePwd,*IndPulse);
		ul_Qty			= *pPulses;
		*pPulses		= 0;
	} 
	else
	{
		(*IndPulse)		= pFilePwdSt->l_Pulse_Count - 1;
		set_pwd_index(pPwdIndex,*IndFilePwdLst,*IndFilePwd,*IndPulse);
		ul_Qty			= ul_Pulses;
		*pPulses		-= ul_Pulses;
	}
	return ul_Qty;
}

__File_Pwd_List_St* Pulse_project::new_FilePwdListSt(unsigned_long IndFilePwdLst)
{
//	__File_Pwd_List_St* pFilePwdList = NULL;//= new __File_Pwd_List_St;
// 	pFilePwdList->p_st_FilePwdArray	= 0;
// 	pFilePwdList->ul_PulseCount		= 0;
// 	pFilePwdList->us_ListCount		= 0;
// 	pFilePwdList->p_PathName		= NULL;
// 	pFilePwdList->p_st_FileList		= NULL;
// 	pFilePwdList->p_st_FilePwdArray	= NULL;

	// _DestroyFilePwdList(mProject.pFilePwdListSt);
	// set_filePwdListSt(pFilePwdList,IndFilePwdLst);
	mProject.IndexFilePwdLstSt		= IndFilePwdLst;
	read_filePwdListSt(IndFilePwdLst);
	//pFilePwdList->p_st_FileList		= 0;
	return mProject.pFilePwdListSt;
}
				
bool Pulse_project::Next_FilePwdSt	(__File_Pwd_List_St* pFilePwdList,
									 unsigned_long *pIndFilePwdLst,
									 unsigned_long *pIndFilePwd,
									 unsigned_long *pIndPulse)
{
	bool result					= true;
	//__File_Pwd_St*	pFilePwdSt	= pFilePwdList->p_st_FilePwdArray+*pIndFilePwd;
	
	*pIndPulse					= 0;
	(*pIndFilePwd)				++;
	if(*pIndFilePwd >= pFilePwdList->us_ListCount)
	{
		*pIndFilePwd			= 0;
		(*pIndFilePwdLst)			++;
		if (*pIndFilePwdLst >= mProject.FilePwdList_Count)
		{
			result				= false;
		}
		else
		{
			//_Destroy(pFilePwdList);
			set_filePwdListSt(pFilePwdList,*pIndFilePwdLst);
			read_filePwdListSt(*pIndFilePwdLst);
		}
	}
	return result;
}
						 
void Pulse_project::Next_Pwd(__File_Pwd_List_St* pFilePwdList,
							 unsigned_long *pIndFilePwdLst,
							 unsigned_long *pIndFilePwd,
							 unsigned_long *pIndPulse)
{
	__File_Pwd_St*	pFilePwdSt	= pFilePwdList->p_st_FilePwdArray+*pIndFilePwd;

	(*pIndPulse)					++;
	if (*pIndPulse >= pFilePwdSt->l_Pulse_Count)
	{
		Next_FilePwdSt(pFilePwdList,pIndFilePwdLst,pIndFilePwd,pIndPulse);
	} 
}

bool	Pulse_project::save_FileLstSt_op(__File_List_St* pFileLstSt,FILE* pFile)
{
	if(write_in_file(pFileLstSt,sizeof(__File_List_St),1,pFile)==false)
	{
		return false;
	}
	if(write_in_file(pFileLstSt->p_ch_FileList,sizeof(__File_List_St),1,pFile)==false)
	{
		return false;
	}
	if(write_in_file(pFileLstSt->p_ch_NamesList,sizeof(char),pFileLstSt->ul_NameTableSize,pFile)==false)
	{
		return false;
	}
	if(write_in_file(pFileLstSt->p_us_NamesLenList,sizeof(long),pFileLstSt->us_ListCount,pFile)==false)
	{
		return false;
	}
	if(write_in_file(pFileLstSt->p_ch_Path,sizeof(char),strlen(pFileLstSt->p_ch_Path),pFile)==false)
	{
		return false;
	}
	return true;
}

bool	Pulse_project::read_FileLstSt_op(__File_List_St* pFileLstSt,FILE* pFile)
{
	if(read_in_file(pFileLstSt,sizeof(__File_List_St),1,pFile)==false)
	{
		return false;
	}
	if(write_in_file(pFileLstSt->p_ch_FileList,sizeof(__File_List_St),1,pFile)==false)
	{
		return false;
	}
	if(write_in_file(pFileLstSt->p_ch_NamesList,sizeof(char),pFileLstSt->ul_NameTableSize,pFile)==false)
	{
		return false;
	}
	if(write_in_file(pFileLstSt->p_us_NamesLenList,sizeof(long),pFileLstSt->us_ListCount,pFile)==false)
	{
		return false;
	}
	if(write_in_file(pFileLstSt->p_ch_Path,sizeof(char),strlen(pFileLstSt->p_ch_Path),pFile)==false)
	{
		return false;
	}
	return true;
}

bool	Pulse_project::save_filePwdListSt_op(void)
{
	FILE*				pFile;
	char				FileName[260];
	__File_Pwd_List_St*	pFilePwdLstSt;
	__File_Pwd_St*		pFilePwdSt;
	__Pwd_St*			pPwdSt;
	__Pwd_NF_St*		pPwdNFSt;
	__File_List_St*		pFileLstSt;
	char*				pPulseFileName;
	//long				Index;
	
	pFilePwdLstSt		= mProject.pFilePwdListSt;
	pFilePwdSt			= pFilePwdLstSt->p_st_FilePwdArray;
	
	get_fileNameFilePwdLstSt(FileName);
	pFile		= fopen(FileName,"wb");
	if (pFile==NULL)
	{
		return false;
	}
	if(write_in_file(pFilePwdLstSt,sizeof(__File_Pwd_List_St),1,pFile)==false)
	{
		return false;
	}
	if(write_in_file(pFilePwdSt,sizeof(__File_Pwd_St),pFilePwdLstSt->us_ListCount,pFile)==false)
	{
		return false;
	}
	pPwdSt			= pFilePwdSt->p_st_Pwd;
	pPwdNFSt		= pFilePwdSt->p_st_Pwd_NewFields;
	pPulseFileName	= pFilePwdSt->p_ch_FileName;
	pFileLstSt		= pFilePwdLstSt->p_st_FileList;
	if(write_in_file(pPwdSt,sizeof(__Pwd_St),pFilePwdLstSt->ul_PulseCount,pFile)==false)
	{
		return false;
	}
	if(write_in_file(pPwdNFSt,sizeof(__Pwd_NF_St),pFilePwdLstSt->ul_PulseCount,pFile)==false)
	{
		return false;
	}
	/*if(write_in_file(pPulseFileName,sizeof(char),pFilePwdSt->us_FileName_length+1,pFile)==false)
	{
		return false;
	}*/
	if(write_in_file(pFileLstSt,sizeof(char),pFilePwdSt->us_FileName_length+1,pFile)==false)
	{
		return false;
	}
	fclose(pFile);
	return true;
}

