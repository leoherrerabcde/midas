/*
 * Pulseproject.h
 *
 *  Created on: Oct 5, 2012
 *      Author: lherrera
 */

#ifndef PULSEPROJECT_H_
#define PULSEPROJECT_H_

#include "pulse_conv_struct_define.h"
#include "PoolMemory.h"
#include "BookMark.h"
#include <stdio.h>
#include <list>
using namespace std;

class Pulse_project {
public:
	Pulse_project();
	virtual ~Pulse_project();

	void CreateNewProjec();

	void	set_filePwdListSt	(__File_Pwd_List_St* pFilePwdLstSt, long Index)
	{ mProject.pFilePwdListSt	= pFilePwdLstSt; mProject.IndexFilePwdLstSt=Index;}	
	bool	save_filePwdListSt	(void);
	bool	save_filePwdListSt_op	(void);
	bool	save_FileLstSt_op	(__File_List_St* pFileLstSt,FILE* pFile);
	bool	read_file			(long index);
	bool	read_filePwdListSt	(long index);
	bool	read_FileLstSt_op	(__File_List_St* pFileLstSt,FILE* pFile);
	void	setMissionPath		(__FileNameSt* pPath);
	void	setWrkSpcPath		(__FileNameSt* pPath);
	void	setOutputPath		(__FileNameSt* pPath);
	void	getMissionName		(char* lvName);
	void	setIntervalPerSheet	(double Interval);
	void	setSheetsPerXls		(long SheetCount);
	void	setPulsesPerSheet	(long PulseCount);
	void	create_workspace	(void);
	void	create_workspace_byTime	(void);
	void	create_workspace_byPulse(void);
	void	create_workspace_byInterval(void);
	void	setColumnFormat		(long* ColumnOrder, long* ColumnEnable);
	bool	SaveWorkSpace		(void);
	bool	LoadWorkSpace		(void);
	bool	SaveColumnFormat	(void);
	bool	SaveErrPntLst		(void);
	bool	SaveErrPntLst		(unsigned_long Index);
	bool	LoadErrPntLst		(unsigned_long Numb);
unsigned_long	LoadErrFileCount(void);
	bool	SaveErrPntLstNew	(void);
	bool	LoadErrPntLstNew	(void);
	void	Set_ErrorPtrLst		(list<__PwdPointerSt>* pErrPntLst);
	void	Set_ErrorPtrLstNew	(__PwdPointerSt* pErrPntLst,unsigned_long ulErrPrtCnt);
	void	ClearErrorFileCount (void) {m_Error_File_Count=0;};

	__ProjectSt		mProject;

// 	long	GetPwdCount(long IndexSpread , 
// 						long IndexSheet ,
// 						long Index);
	long	GetSheetCount (long IndexSpread){
				return mProject.pProjectFile->pSpreadFileArray[IndexSpread].us_WorkSheetCount;}
	long	GetSpreadFileCount(void){
				if (mProject.pProjectFile!=NULL) return mProject.pProjectFile->us_SpreadFileCount;
				else return 0;};
	long	GetPulseCount(long IndexSpread , 
						 long IndexSheet);
	void	GetSheetInfo(long IndexSpread , 
						 long IndexSheet ,
						 long * PulseQty ,
						 double* TimeIni ,
						 double* TimeEnd);
	void	GetSpreadinfo(long IndexSpread , 
						long * PulseQty ,
						double* TimeIni ,
						 double* TimeEnd);
	void	CreateSheet (long IndexSpread,long IndexSheet);
	void	SaveSpreadSheet (long IndexSpread);
	void	GetPwd		(long IndexSpread,
						 long IndexSheet,
						 long IndexPulse,
						 double* pPwd);
	void	get_FileName(long IndexSpread , 
						long IndexSheet ,
						char* lvStr);
	
	long	get_FilePwdList_Count(void){return mProject.FilePwdList_Count;};
	void	set_FilePwdList_Count(long lCount){mProject.FilePwdList_Count = lCount;};

	__File_Pwd_List_St*	Create_SpreadFile(long IndexSpread);
	bool	Create_SpreadFile(__File_Pwd_List_St* pFilePwdLstSt,long IndexSpread);
	void	Destroy(__File_Pwd_List_St* pFilePwdLstSt){_DestroyFilePwdList(pFilePwdLstSt);};
	void	DestroyAll	(void);
	void	DestroyWorkSpace(void);
	void	DestroyErrPntLst(void);
	void	UnSetErrPntLst	(void);

	void    Set_Index_FilePwdList(unsigned_long ulValue){m_Index_FilePwdList=ulValue;};
	unsigned_long   Get_Index_FilePwdList(void){return m_Index_FilePwdList;};

	void    Set_Index_FilePulse(unsigned_long ulValue){m_Index_FilePulse=ulValue;};
	unsigned_long   Get_Index_FilePulse(void){return m_Index_FilePulse;};
	
	void    Set_FileCount(unsigned_long ulValue){m_FileCount=ulValue;};
	unsigned_long   Get_FileCount(void){return m_FileCount;};
	
	void    Set_MaxPulses(unsigned_long ulValue){m_MaxPulses=ulValue;};
	unsigned_long   Get_MaxPulses(void){return m_MaxPulses;};
	
	void    Set_ImportFileDone(void){m_ImportFileDone=1;};
	void    Clear_ImportFileDone(void){m_ImportFileDone=0;};
	unsigned_long	Get_ImportFileDone(void){return m_ImportFileDone;};
	
	void    Set_WrkSpcDone(void){m_WrkSpcDone=true;};
	void    Clear_WrkSpcDone(void){m_WrkSpcDone=true;};
	bool	Get_WrkSpcDone(void){return m_WrkSpcDone;};
	unsigned_long	Get_FileMaxPulses(__File_List_St* pFileLstSt,unsigned_long lvFilelstCount);

	CBookMark		m_cBookMark;

	CPoolMemory<__File_Pwd_List_St>*	m_PoolFilePwdLst;
	CPoolMemory<__File_Pwd_St>*			m_PoolFilePwd;
	CPoolMemory<__Pwd_NF_St>*			m_PoolPwdNF;
	CPoolMemory<__Pwd_St>*				m_PoolPwd;
	CPoolMemory<__PwdPointerSt>*		m_PoolPwdErr;

private:
	void	cpy_file_name				(__FileNameSt* Dst, __FileNameSt* Src);
	void	get_fileNameFilePwdLstSt	(char* pFileName);
	void	get_fileNameFilePwdLstSt	(char* pFileName, long index);
	bool	write_in_file				(void * pSource,
										 size_t size, 
										 size_t count, 
										 FILE* pFile);
	bool	read_in_file				(void * pDst,
										 size_t size, 
										 size_t count, 
										 FILE* pFile);
	void	DestroyFileNameSt			(__FileNameSt* pFileNameSt);
	void	set_pwd_index				(__PwdIndex* pPwdIndex,
										 unsigned_long IndFilePwdLst,
										 unsigned_long IndFilePwd,
										 unsigned_long IndPulse);
	void	PwdIndex_Cpy	(__PwdIndex* pPwdIndexDst,__PwdIndex* pPwdIndexSrc);
	void	Add_Index_Ini	(list<__PwdIndex*>* pListIndex,
							 __PwdIndex* pPwdIndex);
	void	Add_Index_Ini	(list<__PwdIndex*>* pListIndex,
							 unsigned_long IndFilePwdLst,
							 unsigned_long IndFilePwd,
							 unsigned_long IndPulse);
	void	Add_Index_Ini	(list<__PwdIndex*>* pListIndex,
							 __PwdIndex* pPwdIndex,
							 unsigned_long IndFilePwdLst,
							 unsigned_long IndFilePwd,
							 unsigned_long IndPulse);
	void	Add_Index_End	(list<__PwdIndex*>* pListIndex,
							 unsigned_long IndFilePwdLst,
							 unsigned_long IndFilePwd,
							 unsigned_long IndPulse);
	unsigned_long	Add_Pulses_To_Index	(__PwdIndex* pPwdIndex,
								 __File_Pwd_List_St* pFilePwdList,
								 unsigned_long* IndFilePwdLst,
								 unsigned_long* IndFilePwd,
								 unsigned_long* IndPulse,
								 unsigned_long* pPulses);
	unsigned_long	Add_Pulses_To_Index	(__PwdIndex* pPwdIndex,
										__File_Pwd_List_St* pFilePwdList,
										unsigned_long* IndFilePwdLst,
										unsigned_long* IndFilePwd,
										unsigned_long* IndPulse,
										double* p_dTime_ms,
										double d_TimeEnd_ms);
	bool	Next_FilePwdSt	(__File_Pwd_List_St* pFilePwdList,
							 unsigned_long *pIndFilePwdLst,
							 unsigned_long *pIndFilePwd,
							 unsigned_long *pIndPulse);
	void	Next_Pwd		(__File_Pwd_List_St* pFilePwdList,
							 unsigned_long *pIndFilePwdLst,
							 unsigned_long *pIndFilePwd,
							 unsigned_long *pIndPulse);
	void	Init_PwdIndex	(__PwdIndex* pPwdIndex);
	double	TimeNextPulse	(__File_Pwd_List_St* pFilePwdList,
							unsigned_long IndFilePwdLst,
							unsigned_long IndFilePwd,
							unsigned_long IndPulse);

	__WorkSheetBounds*	new_work_sheet_bounds	(list<__PwdIndex*>* pListIndex);
	__SpreadFile*		new_SpreadFile(list<__WorkSheetBounds*>* pWorkSheetBounds);
	__SpreadFileList*	new_SpreadFileList(list<__SpreadFile*>*pSpreadFileList);
	__File_Pwd_List_St* new_FilePwdListSt(unsigned_long IndFilePwdLst);
	__File_Pwd_List_St* new_FilePwdListSt(void);
	__File_Pwd_St*		new_FilePwdSt(unsigned_long ulQty);

	void	DestroyListIndex	(list<__PwdIndex*>* pListIndex);
	void	DestroyListWorkSheetBounds	(list<__WorkSheetBounds*>* pWorkSheetBounds);
	void	DestroyListSpreadFile(list<__SpreadFile*>*pSpreadFileList);

	void	_Destroy(__File_Pwd_List_St* pFilePwdLstSt){_DestroyFilePwdList(pFilePwdLstSt);};
	void	_Destroy(__WorkSheetBounds* pWorkSheetBounds){DestroyWorkSheetBounds(pWorkSheetBounds);};
	void	_Destroy(__SpreadFile* pSpreadFile){DestroySpreadFile(pSpreadFile);};
	void	_Destroy(__SpreadFileList* pSpreadFileList){DestroySpreadFileList(pSpreadFileList);};
	void	_Destroy(char* p){__DestroyArray(p);};
	//void	_Destroy(__File_List_St* p){__Destroy(p);};
	void	_Destroy(__File_Pwd_St* p){_DestroyFilePwdSt(p);};
	void	_Destroy(__File_List_St* p){_DestroyFileList(p);};

	void	_DestroyField(__File_Pwd_St* p){_DestroyFieldFilePwdSt(p);};
	void	_DestroyField(__WorkSheetBounds* p){_DestroyFieldWorkSheetBounds(p);};
	void	_DestroyField(__SpreadFile* p){_DestroyFieldSpreadFile(p);};

	void	_DestroyFieldFilePwdSt(__File_Pwd_St* pFilePwdSt);
	void	_DestroyFieldWorkSheetBounds(__WorkSheetBounds* pWorkSheetBounds);
	void	_DestroyFieldSpreadFile(__SpreadFile* pSpreadFile);

	void	DestroyWorkSheetBounds(__WorkSheetBounds* pWorkSheetBounds);
	void	DestroyWorkSheetBoundsArray(__WorkSheetBounds* pWorkSheetBounds,long lCount);
	void	DestroySpreadFile(__SpreadFile* pSpreadFile);
	void	DestroySpreadFileList(__SpreadFileList* pSpreadFileList);
	void	_DestroyFilePwdList(__File_Pwd_List_St* pFilePwdLstSt);
	void	_DestroyFilePwdSt(__File_Pwd_St* pFilePwdSt);
	void	_DestroyFileList(__File_List_St* pFileListSt);

	void	__DestroyArray(void* p);
	void	__Destroy(void* p);
	template <class T> void	__DestroyList(list<T*>*pList);

	//void	DestroyFilePwdListSt();

	__WorkSheetBounds*	_Get_WorkSheetBound(long IndexSpread,long IndexSheet);
	__SpreadFile*		_Get_SpreadFile(long IndexSpread);
	__File_Pwd_List_St* _Get_FilePwdList(unsigned_long IndFilePwdLst);
/*	__File_Pwd_St*		_Get_FilePwd(unsigned_long IndFilePwd);
	__File_Pwd_St*		_Get_FilePwd(unsigned_long IndFilePwdLst,unsigned_long IndFilePwd);
	__File_Pwd_St*		_Get_FilePwd(long IndexSpread,long IndexSheet,long IndexPulse);*/
	__PwdIndex*			_Get_PwdIndex(__WorkSheetBounds* pWorkSheetBounds,long IndexPulse);
	void	_Get_Sheet(long IndexSpread,long IndexSheet);
	void	_Get_Sheet(__File_Pwd_St* pFilePwd,long IndexSpread,long IndexSheet);

	void	_AddInfoTopPwd (__PwdIndex* pIndex,__File_Pwd_List_St* pFilePwdList);
	void	_cpyIndexData(__Pwd_St* pPwd,__Pwd_NF_St* pPwdNf,__PwdIndex* pIndex);
	void	_IndexToPointers(__PwdIndex* pIndex,
					__File_Pwd_St **pFilePwd,
					__Pwd_NF_St** pPwdNf,
					__Pwd_St** pPwd);
	void	_SaveInfo	(__Ptd_St* pPtd_Dst,__Ptd_St* pPtd_Src,bool* pFlag);
	void	_SaveInfo	(__Ptd_St* pPtd_Dst,__Ptd_St* pPtd_Src){_SaveInfo(pPtd_Dst,pPtd_Src,NULL);};
	void	Save_SpreadFileList(__SpreadFileList *pSpreadFileLst);
	unsigned_long	Get_FilePulses(char* FileName);

	void	SetFileNumber(char *FileNumber, unsigned_long Index);

	unsigned_long	m_Index_FilePwdList;
	unsigned_long	m_Index_FilePulse;
	unsigned_long	m_FileCount;
	unsigned_long	m_ImportFileDone;
	bool			m_WrkSpcDone;
	unsigned_long	m_MaxPulses;
	unsigned_long	m_Error_File_Count;
};

#endif /* PULSEPROJECT_H_ */