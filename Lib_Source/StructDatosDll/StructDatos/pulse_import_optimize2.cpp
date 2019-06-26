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

void pulse_import::import_file_list_op(__file_list_st *fileListStruct,
									__Ptd_St * p_LastPreviousTime,
									__File_Pwd_St* filePwdStructPrev,
									__File_Pwd_List_St	*filePwdLstStruct,
									long* pFileCounter,
									unsigned_long* pCountPulse,
									double* pLastPostDToaus)
{
	__File_Pwd_St*			filePwdStruct;
	__Ptd_St				*ptdStPrev = p_LastPreviousTime;
	unsigned long			i;
	double					d_FileTime_s;
	__Pwd_NF_St*			pPwdNfSt;
	__Pwd_NF_St*			pPwdNfArray;
	__Pwd_St*				pPwdArray;
	__Pwd_NF_St*			pPwdNf;
	__Pwd_St*				pPwd;
	
	filePwdStruct							= filePwdLstStruct->p_st_FilePwdArray;
	m_ErrPrtCnt								= 0;
	m_PoolPwdErr->Free(m_ErrorPtrLstPt);

	pPwdArray								= m_PoolPwd->Alloc(filePwdLstStruct->ul_PulseCount);
	pPwdNfArray								= m_PoolPwdNF->Alloc(filePwdLstStruct->ul_PulseCount);
	m_ErrorPtrLstPt							= m_PoolPwdErr->Alloc(filePwdLstStruct->ul_PulseCount);
	m_pErrPrt								= m_ErrorPtrLstPt;
	pPwdNf									= pPwdNfArray;
	pPwd									= pPwdArray;

	for(i = 0; i < fileListStruct->us_ListCount; i++)
	{
		filePwdStruct->s_WrapAround_Counter	= 0;
		filePwdStruct->uc_OverFlow_Flag		= 0;
		//assign_file_name					(filePwdStruct,fileListStruct,i);
		filePwdStruct->p_st_Pwd				= pPwd;
		filePwdStruct->p_st_Pwd_NewFields	= pPwdNf;
		//read_file							(filePwdStruct, filePwdLstStruct->p_PathName);
		read_file_op						(filePwdStruct, filePwdLstStruct->p_PathName);
		d_FileTime_s						= get_time_file_s(filePwdStruct);
		setStartTime(d_FileTime_s);
		get_first_ptd						(filePwdStruct, d_FileTime_s, ptdStPrev);
		ptdStPrev							= &(filePwdStruct->st_Last_Ptd);
		get_pwd_nf							(filePwdStruct,filePwdStructPrev,*pFileCounter);
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
}

void pulse_import::DestroyFilePwdLstSt_op(__File_Pwd_List_St* filePwdLstStruct)
{
	if (filePwdLstStruct!=NULL)
	{
		_DestroyArraySt(filePwdLstStruct->p_st_FilePwdArray,filePwdLstStruct->us_ListCount);
	}
}

__Pwd_St * pulse_import::read_file_op(__File_Pwd_St *p_st_File_Pwd, char *p_ch_Pulse_Path)
{
	__Pwd_St 		*p_st_Pwd = NULL;
	FILE			*pFile;
	//long			l_FileSize;
	long			l_PlsQty;
	long			l_PlsRead;
	char 			ch_FileName[512];
	
	strcpy(ch_FileName,p_ch_Pulse_Path);
	strcat(ch_FileName, "\\");
	strcat(ch_FileName,p_st_File_Pwd->p_ch_FileName);
	
	pFile               = fopen(ch_FileName,"rb");
	if (pFile == NULL)
	{
		return NULL;
	}
	p_st_Pwd 	= p_st_File_Pwd->p_st_Pwd;
	l_PlsQty	= p_st_File_Pwd->l_Pulse_Count;
	l_PlsRead 	= fread(p_st_Pwd, (size_t)sizeof(__Pwd_St), (size_t)l_PlsQty, pFile);
	fclose(pFile);
	if (l_PlsRead != l_PlsQty)
	{
		return NULL;
	}
	return p_st_Pwd;
}

__Pwd_NF_St	* pulse_import::get_pwd_nf_op(__file_pwd_st * p_st_FilePwd, 
									   __file_pwd_st * p_st_FilePwdPrev,
									   long ul_FileNumber)
{
	__Pwd_NF_St 		* p_st_PwdNF;
	__Pwd_NF_St 		* p_st_PwdNFPrev = NULL;
	__Pwd_St			* p_st_Pwd;
	__Pwd_St			* p_st_Pwd_Prev = NULL;
	__Ptd_St 			* p_st_Ptd;
	__Ptd_St 			* p_st_Ptd_Prev = NULL;
	long				i;
	double				d_Rel_S_Toa_ms = 0;
	unsigned long		ul_Interval;
	double				d_Interval_ms;
	double				d_PulseDTimeInterval_s;

	p_st_PwdNF	= p_st_FilePwd->p_st_Pwd_NewFields;

	p_st_Ptd							= &(p_st_FilePwd->st_First_Ptd);
	p_st_Pwd							= p_st_FilePwd->p_st_Pwd;

	if (p_st_FilePwdPrev != NULL)
	{
		i									= p_st_FilePwdPrev->l_Pulse_Count-1;
		p_st_Pwd_Prev						= p_st_FilePwdPrev->p_st_Pwd + i;
		p_st_PwdNFPrev						= p_st_FilePwdPrev->p_st_Pwd_NewFields + i;
		p_st_Ptd_Prev						= &(p_st_FilePwdPrev->st_Last_Ptd);
	}

	for(i = 0 ; i < p_st_FilePwd->l_Pulse_Count ; i++)
	{
		p_st_PwdNF->uc_Toa_Error	= 0;
		if(!i)
		{
			p_st_PwdNF->uc_RollOver			= p_st_FilePwd->uc_OverFlow_Flag;
			p_st_PwdNF->uc_WrapAround		= p_st_FilePwd->s_WrapAround_Counter;
			p_st_PwdNF->d_FileTime_s		= p_st_FilePwd->st_First_Ptd.d_FileTime_s;
			p_st_PwdNF->uc_Toa_Error		= 0;
			if (p_st_Ptd_Prev != NULL)
			{
				d_Interval_ms				= p_st_Ptd->d_D_Toa_ms;
				d_Rel_S_Toa_ms				= p_st_Ptd->d_Time_ms;
				p_st_PwdNF->d_post_d_Toa_us	= d_Interval_ms * MSEC_TO_USEC_FACTOR;
				p_st_PwdNF->d_post_d_Toa_ms	= d_Interval_ms;
				p_st_PwdNF->d_Rel_Toa_ms	= d_Rel_S_Toa_ms;
				p_st_PwdNFPrev->d_pre__d_Toa_us	= p_st_PwdNF->d_post_d_Toa_us;
				p_st_PwdNF->d_Abs_Toa_s		= p_st_PwdNFPrev->d_Abs_Toa_s + d_Interval_ms / DTOAS_TO_SEC_FACTOR;
			}
			else
			{
				p_st_PwdNF->d_pre__d_Toa_us	= 0;
				p_st_PwdNF->d_post_d_Toa_us	= 0;
				p_st_PwdNF->d_post_d_Toa_ms	= 0;
				p_st_PwdNF->d_Rel_Toa_ms	= 0;
				p_st_PwdNF->d_Rel_S_Toa_ms	= 0;
				p_st_PwdNF->d_Abs_Toa_s		= m_d_Start_Time_s;
				d_Interval_ms				= p_st_Ptd->d_D_Toa_ms;
			}
			p_st_PwdNF->d_Rel_S_Toa_ms		= d_Rel_S_Toa_ms;
			p_st_PwdNF->d_Date_s			= m_d_Start_Time_s + p_st_PwdNF->d_Rel_Toa_ms / DTOAS_TO_SEC_FACTOR;
			d_PulseDTimeInterval_s			= (p_st_PwdNF->d_Date_s - p_st_PwdNF->d_FileTime_s);
			if (d_PulseDTimeInterval_s > DTOA_INTERVAL_SEC_TOL)
			{
				p_st_PwdNF->uc_Toa_Error	|= FileTime_Different_Exceed_Max_Exception;
			}
		} //if(!i)
		else
		{
			ul_Interval						= p_st_Pwd->ul_Toa - p_st_Pwd_Prev->ul_Toa;
			if (p_st_Pwd->ul_Toa <= p_st_Pwd_Prev->ul_Toa)
			{
				ul_Interval					+= TOA_OVFW_LESS_ONE;
				ul_Interval					++;
				p_st_PwdNF->uc_RollOver		= 1;
			}
			else
			{
				p_st_PwdNF->uc_RollOver		= 0;
			}
			d_Interval_ms					= (double)ul_Interval / TOA_TO_MSEC_FACTOR;
			if (d_Interval_ms > MAX_DTOA_INTERFILE_MAX_MS )
			{
				p_st_PwdNF->uc_Toa_Error	|= DToa_InterFile_Exceed_Max_Exception;
			} 
			d_PulseDTimeInterval_s			= (p_st_PwdNF->d_Date_s - p_st_PwdNF->d_FileTime_s);
			if (d_PulseDTimeInterval_s > DTOA_INTERVAL_SEC_TOL)
			{
				p_st_PwdNF->uc_Toa_Error	|= FileTime_Different_Exceed_Max_Exception;
			}
			d_Rel_S_Toa_ms					+= d_Interval_ms;
			p_st_PwdNF->uc_WrapAround		= 0;
			p_st_PwdNF->d_post_d_Toa_us		= (double)ul_Interval / TOA_TO_USEC_FACTOR;
			p_st_PwdNF->d_post_d_Toa_ms		= d_Interval_ms;
			p_st_PwdNF->d_Rel_Toa_ms		= p_st_PwdNFPrev->d_Rel_Toa_ms + d_Interval_ms;
			p_st_PwdNF->d_Rel_S_Toa_ms		= d_Rel_S_Toa_ms;
			p_st_PwdNF->d_Abs_Toa_s			= p_st_PwdNFPrev->d_Abs_Toa_s + d_Interval_ms / DTOAS_TO_SEC_FACTOR;
			p_st_PwdNF->d_Date_s			= m_d_Start_Time_s + p_st_PwdNF->d_Rel_Toa_ms / DTOAS_TO_SEC_FACTOR;
			p_st_PwdNF->d_FileTime_s		= p_st_FilePwd->st_First_Ptd.d_FileTime_s;
		} // if(!i) else
		p_st_PwdNF->uc_PulseDetail			= 0;
		p_st_PwdNF->ul_Index				= m_ul_Correl_Index++;
		p_st_PwdNF->ul_Rel_Index			= m_ul_Relati_Index++;

		if (p_st_PwdNFPrev != NULL)
		{
			p_st_PwdNFPrev->d_pre__d_Toa_us	= p_st_PwdNF->d_post_d_Toa_us;
		}
		if (m_p_d_pre__d_Toa_us!=NULL)
		{
			if (*m_p_d_pre__d_Toa_us != p_st_PwdNF->d_post_d_Toa_us	)
			{
				*m_p_d_pre__d_Toa_us		= p_st_PwdNF->d_post_d_Toa_us;
			}
		}
		m_p_d_pre__d_Toa_us					= &(p_st_PwdNF->d_pre__d_Toa_us);
		p_st_PwdNF->us_FileNumber			= ul_FileNumber+1;
		// Verify Errors
		VerifyPwdErrors(p_st_Pwd,p_st_PwdNF,p_st_PwdNFPrev,i);
		VerifyPwdErrors_op(p_st_Pwd,p_st_PwdNF,p_st_PwdNFPrev,i);
		// End Verify Errors
		p_st_Pwd_Prev						= p_st_Pwd++;
		p_st_PwdNFPrev						= p_st_PwdNF++;
	}  // for(i = 0 ; i < p_st_FilePwd->l_Pulse_Count ; i++)

	return p_st_FilePwd->p_st_Pwd_NewFields;
}

unsigned char pulse_import::VerifyPwdErrors_op(__Pwd_St* p_st_Pwd,
											__Pwd_NF_St* p_st_PwdNF,
											__Pwd_NF_St* p_st_PwdNF_Previous,
											unsigned_long ul_IndexPwd)
{
	__ErrorUnionSt			sCodeError;
	static unsigned char	ssLastError = 0;
	
	sCodeError.ErrorCode = 0;
	
	sCodeError.ErrorFlags.Frec_Err	= VerifyFrec(p_st_Pwd->ul_IFM);
	sCodeError.ErrorFlags.Amp_Err	= VerifyAmplitud(p_st_Pwd->si_Amplitud);
	sCodeError.ErrorFlags.Pw_Err	= VerifyPW(p_st_Pwd->us_PulseWidth);
	
	if (p_st_PwdNF_Previous!=NULL)
	{
		sCodeError.ErrorFlags.Neg_DToa		= VerifyDToaError	(p_st_PwdNF->d_post_d_Toa_us);
		sCodeError.ErrorFlags.Rel_Toa_Err	= VerifyRToaError	(p_st_PwdNF->d_Rel_Toa_ms,
			p_st_PwdNF_Previous->d_Rel_Toa_ms);
		sCodeError.ErrorFlags.Abs_Toa_Err	= VerifyNewDate	(p_st_PwdNF->d_Abs_Toa_s,
			p_st_PwdNF_Previous->d_Abs_Toa_s);
		sCodeError.ErrorFlags.FileTimeErr	= VerifyFileDate	(p_st_PwdNF->d_Abs_Toa_s,p_st_PwdNF->d_FileTime_s);
	}
	
	if (ssLastError!=sCodeError.ErrorCode)
	{
		ssLastError=sCodeError.ErrorCode;
	}
	p_st_PwdNF->st_ProcessError.ErrorCode	= sCodeError.ErrorCode;
	
	if (sCodeError.ErrorCode)
	{
		__PwdPointerSt*	p_ErrPrt	= m_pErrPrt++;
		m_ErrPrtCnt					++;
		
		if (p_ErrPrt)
		{
			p_ErrPrt->d_Abs_Toa_s		= p_st_PwdNF->d_Abs_Toa_s;
			p_ErrPrt->d_post_d_Toa_us	= p_st_PwdNF->d_post_d_Toa_us;
			p_ErrPrt->d_Rel_Toa_ms		= p_st_PwdNF->d_Rel_Toa_ms;
			p_ErrPrt->s_Error_Code		= sCodeError.ErrorCode;
			p_ErrPrt->ul_Index			= p_st_PwdNF->ul_Index;
			p_ErrPrt->ul_Toa			= p_st_Pwd->ul_Toa;
			p_ErrPrt->ul_IndexFile		= p_st_PwdNF->us_FileNumber;
			p_ErrPrt->ul_IndexPwd		= ul_IndexPwd;
			
			//m_ErrorPtrLst.push_back(st_ErrPrt);
		}
	}
	return sCodeError.ErrorCode;
}

