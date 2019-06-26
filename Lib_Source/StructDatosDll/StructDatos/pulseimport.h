/*
 * pulseimport.h
 *
 *  Created on: Sep 16, 2012
 *      Author: lherrera
 */

#ifndef PULSEIMPORT_H_
#define PULSEIMPORT_H_

#include "filelist.h"
#include "pulse_conv_struct_define.h"
#include "Pulseproject.h"
#include <list>
using namespace std;

#define TOA_OVFW_LESS_ONE 	4294967295U
#define OVERFLOW_SEC		42.94967296L
#define OVERFLOW_MSEC		42949.67296L
#define TOA_TO_USEC_FACTOR	100.0L
#define TOA_TO_MSEC_FACTOR	100000.0L
#define MSEC_TO_USEC_FACTOR 1000.0L
#define DTOAS_TO_SEC_FACTOR	1000.0L
#define DTOA_INTERVAL_SEC_TOL	1.0L

#define TIMET_TO_DAY_FACTOR	86400.0L

#include "PoolMemory.h"

class pulse_import {
public:
	pulse_import();
	virtual ~pulse_import();

	__File_Pwd_List_St* import_file(char * p_ch_Pulse_Path);
	void				import_file(__ProjectSt * pProject);
	void				import_file(Pulse_project* pPulseProject);
	void				import_file_v2(Pulse_project* pPulseProject);
	void				import_file_optimize(Pulse_project* pPulseProject);

	char* 				get_pulse_file(unsigned short Index, char * FilePulseName);
	int					get_file_len(unsigned short Index);
	double* 			get_pwd(unsigned short IndexPwdFile, unsigned long PulseIndex, double * Pwd);
	void				cleanStartTime(void){m_d_StartTimeEmpty=true;};

	CPoolMemory<__File_Pwd_List_St>*	m_PoolFilePwdLst;
	CPoolMemory<__File_Pwd_St>*			m_PoolFilePwd;
	CPoolMemory<__Pwd_NF_St>*			m_PoolPwdNF;
	CPoolMemory<__Pwd_St>*				m_PoolPwd;
	CPoolMemory<__PwdPointerSt>*		m_PoolPwdErr;

private:
	// Private Methods
	__File_Pwd_List_St	* import_file_list(__file_list_st *fileListStruct, 
										   __Ptd_St * p_LastPreviousTime,
										   long* pFileCounter);
	void 				import_file_list(__file_list_st *fileListStruct,
										 __Ptd_St * p_LastPreviousTime,
										 __File_Pwd_St* filePwdStructPrev,
										 __File_Pwd_List_St	*pfilePwdLstStruct,
										 long* pFileCounter,
										 unsigned_long* pCountPulse,
										 double* pLastPostDToaus);
	void 	import_file_list_op(__file_list_st *fileListStruct,
								__Ptd_St * p_LastPreviousTime,
								__File_Pwd_St* filePwdStructPrev,
								__File_Pwd_List_St	*pfilePwdLstStruct,
								long* pFileCounter,
								unsigned_long* pCountPulse,
								double* pLastPostDToaus);
	void pulse_import::DestroyFilePwdLstSt_op(__File_Pwd_List_St* filePwdLstStruct);

	__File_Pwd_List_St*	SetFileListStArray	(Pulse_project* pPulseProject,
								cFilesList* p_cFileList,
								char *pPath);
	void	SetFilePwdLstSt		(__File_Pwd_List_St* pFilePwdLstStNew,__File_List_St* pFileLstSt);
	void	SetFilePwd			(__File_Pwd_St* pFilePwdSt,char *pFileName,char* p_ch_Path);

	unsigned_long	Get_FilePulses	(char* pPath,char* pFileName);

	__File_List_St		* get_file_list(char * p_ch_Pulse_Path);
	__Pwd_St 			* read_file(__File_Pwd_St *p_st_File_Pwd, char *p_ch_Pulse_Path);
	__Pwd_St			* read_file_op(__File_Pwd_St *p_st_File_Pwd, char *p_ch_Pulse_Path);
	double				get_time_file_s(__File_Pwd_St *p_st_File_Pwd);
	double				get_time_file_s(char *p_ch_FileName, unsigned short us_FileName_Length);
	__Ptd_St			* get_last_ptd(__file_pwd_st * p_st_File_Pwd);
	__Ptd_St			* get_first_ptd(__file_pwd_st * p_st_FilePwd, double d_FileTime_s, __Ptd_St * p_LastPreviousTime);
	__Pwd_NF_St			* get_pwd_nf(__file_pwd_st * p_st_FilePwd, 
									 __file_pwd_st * p_st_FilePwdPrev,
									 long ul_FileNumber);
	__Pwd_NF_St			* get_pwd_nf_op(__file_pwd_st * p_st_FilePwd, 
										__file_pwd_st * p_st_FilePwdPrev,
										long ul_FileNumber);
	void				assign_file_name(__file_pwd_st * p_st_FilePwd,
										 __File_List_St * p_File_List_St,
										 long Index);
	void				assign_path_name(__File_Pwd_List_St* p_FilePwdLstSt,
										 __File_List_St * p_File_List_St);
	void 				destroy_file_list(__File_List_St * p_File_List_St);
	void 				destroy_File_Pwd_St(__File_Pwd_List_St *p_File_Pwd_List_St);
	void				assign_pre_dtoa(__Pwd_NF_St* pPwdNfStSrc,__Pwd_NF_St* pPwdNfDst);
	void				assign_pre_dtoa(__File_Pwd_List_St* p_FilePwdLstStPrev, 
										__File_Pwd_List_St* p_FilePwdLstStNow);
	void				_InitFilePwdLstSt(__File_Pwd_List_St* pFilePwdLst);
	void				InitFilePwdSt(__File_Pwd_St &pFilePwdSt,
									__Pwd_St &pPwdSt,
									__Pwd_NF_St &pPwdNFSt);

	void	_FieldDestroy(__File_Pwd_List_St* pFilePwdLst);
	void	_DestroyArray(__File_Pwd_St* pFilePwdArray);
	void	_DestroyArraySt(__File_Pwd_St* pFilePwdArray,long lvCount);
// 	void	_DestroyArray(char* p) {if(p!=NULL){delete []p;}};
// 	void	_DestroyArray(__Pwd_St* p){if(p!=NULL){delete []p;}};
// 	void	_DestroyArray(__Pwd_NF_St* p){if(p!=NULL){delete []p;}};
	void	_DestroyArray(char* p) 
	{
		if(p!=NULL)
		{delete []p;}
	};
	void	_DestroyArray(__Pwd_St* p)
	{
		m_PoolPwd->Free(p);
	};
	void	_DestroyArray(__Pwd_NF_St* p)
	{
		m_PoolPwdNF->Free(p);
	};

	void	setStartTime(double NewStartTime);
	double	getStartTime(void){return m_d_Start_Time_s;};
	char*	_strcpy(char* str);

	unsigned char VerifyPwdErrors(__Pwd_St* p_st_Pwd,
							__Pwd_NF_St* p_st_PwdNF,
							__Pwd_NF_St* p_st_PwdNF_Previous,
							unsigned_long ul_IndexPwd);
	unsigned char VerifyPwdErrors_op(__Pwd_St* p_st_Pwd,
							__Pwd_NF_St* p_st_PwdNF,
							__Pwd_NF_St* p_st_PwdNF_Previous,
							unsigned_long ul_IndexPwd);
	short	VerifyPW(unsigned short usPW);
	short	VerifyAmplitud(short sAmp);
	short	VerifyFrec(unsigned_long lFrec);
	short	VerifyRange(unsigned_long ulVal, unsigned_long ulMin, unsigned_long ulMax);
	short	VerifyRange(unsigned short ulVal, unsigned short ulMin, unsigned short ulMax);
	short	VerifyRange(long lVal, long lMin, long lMax);
	short	VerifyRange(short lVal, short lMin, short lMax);
	short	VerifyNewDate(double NewDate, double PreviusDate);
	short	VerifyFileDate(double NewDate_seg, double FileDate_seg);
	short	VerifyRToaError(double d_RToa,double d_RToa_Previus);
	short	VerifyDToaError(double dDToa_us);

	//unsigned short		m_us_FileCounter;
	double 				m_d_Start_Time_s;				// FileTime(0)
	bool				m_d_StartTimeEmpty;

	__File_Pwd_List_St 	* m_File_Pwd_List_St_Recent;
	__File_List_St		* m_File_List_St_Recent;
	__ProjectSt			* m_ProjectStruct;

	unsigned long		m_ul_Correl_Index;
	unsigned long		m_ul_Relati_Index;
	double*				m_p_d_pre__d_Toa_us;
	double*				m_p_dbg_d_pre__d_Toa_us;
	long				m_Tick_Ini;
	long				m_Enlased_Tick;
	list<__PwdPointerSt>m_ErrorPtrLst;

	__PwdPointerSt*		m_ErrorPtrLstPt;
	__PwdPointerSt*		m_pErrPrt;
	unsigned_long		m_ErrPrtCnt;
	unsigned_long		m_ErrPrtMaxCnt;
	unsigned_long		m_FileIndex;

#endif /* PULSEIMPORT_H_ */
};
