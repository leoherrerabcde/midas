/*
 * pulseimport.cpp
 *
 *  Created on: Sep 16, 2012
 *      Author: lherrera
 */

#include "pulse_conv_struct_define.h"
#include "pulseimport.h"
#include "filelist.h"
#include "PoolMemory.h"

#include <stdio.h>
#include <stdlib.h>
#include <time.h>
#include <dirent.h>
#include <string.h>
#include <iostream>

pulse_import::pulse_import() {
	// TODO Auto-generated constructor stub
	m_File_Pwd_List_St_Recent	= NULL;
	m_File_List_St_Recent		= NULL;
	m_d_StartTimeEmpty			= false;
	m_ErrorPtrLst.clear();
	m_ErrorPtrLstPt			= NULL;
	m_pErrPrt				= NULL;
	m_ErrPrtCnt				= 0;

	m_PoolFilePwdLst		= NULL;
	m_PoolFilePwd			= NULL;
	m_PoolPwdNF				= NULL;
	m_PoolPwd				= NULL;
}

pulse_import::~pulse_import() {
	// TODO Auto-generated destructor stub
	destroy_file_list(m_File_List_St_Recent);
	destroy_File_Pwd_St(m_File_Pwd_List_St_Recent);
}

void pulse_import::destroy_file_list(__File_List_St * p_File_List_St)
{

}

void pulse_import::destroy_File_Pwd_St(__File_Pwd_List_St *p_File_Pwd_List_St)
{

}

// Obtener la lista de archivos de pulsos contenidos en 'p_ch_Pulse_Path'
__File_List_St * pulse_import::get_file_list(char * p_ch_Pulse_Path)
{
	file_list						class_FileList;

	return class_FileList.get_file_list(p_ch_Pulse_Path);
}

__Pwd_St * pulse_import::read_file(__File_Pwd_St *p_st_File_Pwd, char *p_ch_Pulse_Path)
{
	__Pwd_St 		*p_st_Pwd = NULL;
	FILE			*pFile;
	long			l_FileSize;
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
	fseek( pFile, 0, SEEK_END );
	l_FileSize	=  ftell( pFile );
	rewind(pFile);
	l_PlsQty 	= l_FileSize / sizeof(__Pwd_St);
	p_st_File_Pwd->l_Pulse_Count	= l_PlsQty;
	if (l_FileSize % sizeof(__Pwd_St))
	{
		return NULL;
	}
	//p_st_File_Pwd->p_st_Pwd	= new __Pwd_St[l_PlsQty];
	m_PoolPwd->VerifyAllocated(p_st_File_Pwd->p_st_Pwd);
	p_st_File_Pwd->p_st_Pwd	= m_PoolPwd->Alloc(l_PlsQty);
	p_st_Pwd 	= p_st_File_Pwd->p_st_Pwd;
	l_PlsRead 	= fread(p_st_Pwd, (size_t)sizeof(__Pwd_St), (size_t)l_PlsQty, pFile);
	fclose(pFile);
	if (l_PlsRead != l_PlsQty)
	{
		//delete [] p_st_Pwd;
		m_PoolPwd->Free(p_st_Pwd);
		return NULL;
	}
	return p_st_Pwd;
}

double pulse_import::get_time_file_s(char *p_ch_FileName, unsigned short us_FileName_Length)
{
	double 		d_FileTime_s;
	char		ch_DateTimeFile[20];
	short 		s_OffSet, i, s_NullPos;
	struct tm	st_tm_FileTime;
	time_t		st_time_t_s;

	s_OffSet	= us_FileName_Length - 4 - 19;
	memcpy(ch_DateTimeFile, (p_ch_FileName + s_OffSet), 20);

	s_OffSet	= 0;
	s_NullPos	= 4;
	
	for(i = 0; i < 6 ; i++)
	{
		ch_DateTimeFile[s_NullPos] = '\0';
		switch(i)
		{
		case 0:
			st_tm_FileTime.tm_year	= atoi((ch_DateTimeFile + s_OffSet)) - 1900;
			break;
		case 1:
			st_tm_FileTime.tm_mon 	= atoi((ch_DateTimeFile + s_OffSet)) - 1;
			break;
		case 2:
			st_tm_FileTime.tm_mday 	= atoi((ch_DateTimeFile + s_OffSet));
			break;
		case 3:
			st_tm_FileTime.tm_hour 	= atoi((ch_DateTimeFile + s_OffSet));
			break;
		case 4:
			st_tm_FileTime.tm_min 	= atoi((ch_DateTimeFile + s_OffSet));
			break;
		case 5:
			st_tm_FileTime.tm_sec 	= atoi((ch_DateTimeFile + s_OffSet));
			break;
		}
		if (!i)
		{
			s_OffSet	+= 5;
		}
		else
		{
			s_OffSet	+= 3;
		}
		s_NullPos	+= 3;
	}

	st_time_t_s		= mktime(&st_tm_FileTime);
	d_FileTime_s	= (double) st_time_t_s;

	return d_FileTime_s;
}

double pulse_import::get_time_file_s(__File_Pwd_St *p_st_File_Pwd)
{
	double 		d_FileTime_s;

	d_FileTime_s		= get_time_file_s(p_st_File_Pwd->p_ch_FileName, p_st_File_Pwd->us_FileName_length);

	return d_FileTime_s;
}


__Ptd_St * pulse_import::get_last_ptd(__file_pwd_st * p_st_File_Pwd)
{
	long		l_Index 		= p_st_File_Pwd->l_Pulse_Count-1;
	__Pwd_St	* p_st_Pwd 		= p_st_File_Pwd->p_st_Pwd + l_Index;
	__Pwd_NF_St	* p_st_PwdNF	= p_st_File_Pwd->p_st_Pwd_NewFields + l_Index;

	p_st_File_Pwd->st_Last_Ptd.ul_Toa		= p_st_Pwd->ul_Toa;
	p_st_File_Pwd->st_Last_Ptd.d_D_Toa_ms	= p_st_PwdNF->d_post_d_Toa_ms;
	p_st_File_Pwd->st_Last_Ptd.d_FileTime_s	= p_st_PwdNF->d_Abs_Toa_s;
	p_st_File_Pwd->st_Last_Ptd.d_Time_ms	= p_st_PwdNF->d_Rel_Toa_ms;

	return &(p_st_File_Pwd->st_Last_Ptd);
}

__Ptd_St * pulse_import::get_first_ptd(__file_pwd_st * p_st_FilePwd, 
									   double d_FileTime_s, 
									   __Ptd_St * p_LastPreviousTime)
{
	__Ptd_St 		*p_st_FirstPtd;
	double 			d_IntervalFileTime_s;
	double			d_IntervalDToa_ms;
	//double			d_PulseDTimeInterval_s;
	unsigned long	ul_IntervalToa;
	short			s_WrapAroundCounter;
	bool			b_OverFlowFlag = false;

	p_st_FirstPtd 				= &(p_st_FilePwd->st_First_Ptd);
	p_st_FirstPtd->ul_Toa		= p_st_FilePwd->p_st_Pwd->ul_Toa;
	if (p_LastPreviousTime == NULL)
	{
		p_st_FirstPtd->d_D_Toa_ms	= 0;
		p_st_FirstPtd->d_Time_ms	= 0;
		p_st_FirstPtd->d_FileTime_s	= d_FileTime_s;
	}
	else
	{
		//Getting the interval DTOA
		ul_IntervalToa		= p_st_FirstPtd->ul_Toa - p_LastPreviousTime->ul_Toa;
		if (p_LastPreviousTime->ul_Toa >= p_st_FirstPtd->ul_Toa)
		{
			b_OverFlowFlag		= true;
			ul_IntervalToa		+= TOA_OVFW_LESS_ONE;
			ul_IntervalToa		++;
		}
// 		else
// 		{
// 		}
		d_IntervalDToa_ms			= ((double)ul_IntervalToa)/TOA_TO_MSEC_FACTOR;

		d_IntervalFileTime_s		= d_FileTime_s - p_LastPreviousTime->d_FileTime_s;

		s_WrapAroundCounter		= 0;
		if (d_IntervalFileTime_s >= OVERFLOW_SEC)
		{
			s_WrapAroundCounter	= (short) (d_IntervalFileTime_s / OVERFLOW_SEC);
		}

		if (s_WrapAroundCounter)
		{
			d_IntervalDToa_ms	+= (s_WrapAroundCounter * OVERFLOW_MSEC);
		}
		p_st_FirstPtd->d_FileTime_s	= d_FileTime_s;
		/*d_PulseDTimeInterval_s = d_IntervalDToa_ms / DTOAS_TO_SEC_FACTOR - d_IntervalFileTime_s) 
		if (d_PulseDTimeInterval_s > DTOA_INTERVAL_SEC_TOL)
		{
			
		}*/
		/*while ( (d_PulseDTimeInterval_s = d_IntervalDToa_ms / DTOAS_TO_SEC_FACTOR - d_IntervalFileTime_s) > DTOA_INTERVAL_SEC_TOL)
		{
			if (b_OverFlowFlag)
			{
				b_OverFlowFlag		= false;
				d_IntervalDToa_ms	-= OVERFLOW_MSEC;
			}
			else if (s_WrapAroundCounter)
			{
				s_WrapAroundCounter --;
				d_IntervalDToa_ms	-= OVERFLOW_MSEC;
			}
			else
			{
				break;
			}
		}*/
		p_st_FirstPtd->d_D_Toa_ms	= d_IntervalDToa_ms;
		p_st_FirstPtd->d_postDtoa_us= d_IntervalDToa_ms * MSEC_TO_USEC_FACTOR;
		p_st_FirstPtd->d_Time_ms	= p_LastPreviousTime->d_Time_ms + d_IntervalDToa_ms;
		p_st_FilePwd->uc_OverFlow_Flag	= b_OverFlowFlag;
		p_st_FilePwd->s_WrapAround_Counter	= s_WrapAroundCounter;
	}
	return p_st_FirstPtd;
}

__Pwd_NF_St	* pulse_import::get_pwd_nf(__file_pwd_st * p_st_FilePwd, 
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

	//p_st_PwdNF 	= new __Pwd_NF_St[p_st_FilePwd->l_Pulse_Count];
	p_st_PwdNF 	= m_PoolPwdNF->Alloc(p_st_FilePwd->l_Pulse_Count);

	p_st_FilePwd->p_st_Pwd_NewFields	= p_st_PwdNF;

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
		/*if (ul_FileNumber==20 && i == 0)
		{
			p_st_PwdNF->ul_Index = p_st_PwdNF->ul_Index;
		}
		if (ul_FileNumber==194 && i == 0)
		{
			i=i;
		}
		if (ul_FileNumber==193 && i == p_st_FilePwd->l_Pulse_Count-1)
		{
			i=i;
		}*/
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
				//p_st_PwdNF->d_Abs_Toa_s		= m_d_Start_Time_s + p_st_PwdNF->d_Rel_Toa_ms / DTOAS_TO_SEC_FACTOR;
				//p_st_PwdNF->d_Abs_Toa_s		= p_st_PwdNF->d_Rel_Toa_ms / DTOAS_TO_SEC_FACTOR;
				//p_st_PwdNF->d_Abs_Toa_s		+= m_d_Start_Time_s;
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
			//p_st_PwdNF->d_Date_s			= p_st_FilePwd->st_First_Ptd.d_FileTime_s + p_st_PwdNF->d_Rel_Toa_ms / DTOAS_TO_SEC_FACTOR;
			p_st_PwdNF->d_Date_s			= m_d_Start_Time_s + p_st_PwdNF->d_Rel_Toa_ms / DTOAS_TO_SEC_FACTOR;
			p_st_PwdNF->d_FileTime_s		= p_st_FilePwd->st_First_Ptd.d_FileTime_s;
		} // if(!i) else
		p_st_PwdNF->uc_PulseDetail			= 0;
		//p_st_PwdNF->us_FileNumber			= m_us_FileCounter;
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
		//p_st_PwdNFPrev->us_FileNumber		= ul_FileNumber+1;
		p_st_PwdNF->us_FileNumber			= ul_FileNumber+1;
		// Verify Errors
		VerifyPwdErrors(p_st_Pwd,p_st_PwdNF,p_st_PwdNFPrev,i);
		// End Verify Errors
		p_st_Pwd_Prev						= p_st_Pwd++;
		p_st_PwdNFPrev						= p_st_PwdNF++;
	}  // for(i = 0 ; i < p_st_FilePwd->l_Pulse_Count ; i++)

	/*if (ul_FileNumber==193)
	{
		i=i;
	}
	if (m_p_dbg_d_pre__d_Toa_us==NULL)
	{
		m_p_dbg_d_pre__d_Toa_us	= m_p_d_pre__d_Toa_us;
	}*/
	return p_st_FilePwd->p_st_Pwd_NewFields;
}

__File_Pwd_List_St	* pulse_import::import_file(char * p_ch_Pulse_Path)
{
	__File_List_St 			*fileListStruct;
	__File_Pwd_List_St		*filePwdLstStruct;
	struct __file_pwd_st	*filePwdStruct;
	struct __file_pwd_st	*filePwdStructPrev = NULL;
	__Ptd_St				*ptdStPrev = NULL;
	unsigned long			i;
	double					d_FileTime_s;

	// Getting the all pulses files list at 'p_ch_Pulse_Path'
	destroy_file_list		(m_File_List_St_Recent);
	fileListStruct			= get_file_list(p_ch_Pulse_Path);
	m_File_List_St_Recent	= fileListStruct;

	//filePwdLstStruct		= new __File_Pwd_List_St;
	filePwdLstStruct		= m_PoolFilePwdLst->Alloc();

	filePwdLstStruct->us_ListCount 		= fileListStruct->us_ListCount;;
	filePwdLstStruct->p_PathName 		= _strcpy(p_ch_Pulse_Path);
	filePwdLstStruct->s_PathNameLenght 	= strlen(p_ch_Pulse_Path);
	//filePwdLstStruct->p_st_FilePwdArray = new struct __file_pwd_st[filePwdLstStruct->us_ListCount] ;
	filePwdLstStruct->p_st_FilePwdArray = m_PoolFilePwd->Alloc(filePwdLstStruct->us_ListCount) ;
	filePwdLstStruct->p_st_FileList		= fileListStruct;

	//m_us_FileCounter	= 0;
	filePwdStruct						= filePwdLstStruct->p_st_FilePwdArray;
	filePwdLstStruct->ul_PulseCount		= 0;
	for(i = 0; i < fileListStruct->us_ListCount; i++)
	{
		filePwdStruct->s_WrapAround_Counter	= 0;
		filePwdStruct->uc_OverFlow_Flag		= 0;
		//filePwdStruct->p_ch_FileName		= fileListStruct->p_ch_FileList[i];
		//filePwdStruct->us_FileName_length	= fileListStruct->p_us_NamesLenList[i];
		assign_file_name					(filePwdStruct,fileListStruct,i);
		read_file							(filePwdStruct, p_ch_Pulse_Path);
		d_FileTime_s						= get_time_file_s(filePwdStruct);
		if (!i)
		{
			m_d_Start_Time_s 					= d_FileTime_s;
		}
		get_first_ptd						(filePwdStruct, d_FileTime_s, ptdStPrev);
		ptdStPrev							= &(filePwdStruct->st_Last_Ptd);
		get_pwd_nf							(filePwdStruct,filePwdStructPrev,i);
		get_last_ptd						(filePwdStruct);
		filePwdLstStruct->ul_PulseCount		+= filePwdStruct->l_Pulse_Count;
		//m_us_FileCounter 					++;
		filePwdStructPrev					= filePwdStruct;
		filePwdStruct						++;
	}
	filePwdLstStruct->d_TotalTime_ms		= filePwdStructPrev->st_Last_Ptd.d_Time_ms;
	destroy_File_Pwd_St(m_File_Pwd_List_St_Recent);
	m_File_Pwd_List_St_Recent = filePwdLstStruct;
	return filePwdLstStruct;
}

int	pulse_import::get_file_len(unsigned short Index)
{
	return m_File_List_St_Recent->p_us_NamesLenList[Index];
}

char * pulse_import::get_pulse_file(unsigned short Index, char * FilePulseName)
{
	strcpy(FilePulseName, m_File_List_St_Recent->p_ch_FileList[Index]);

	return FilePulseName;
}

double * pulse_import::get_pwd(unsigned short IndexPwdFile, unsigned long PulseIndex, double * Pwd)
{
	return Pwd;
}

char* pulse_import::_strcpy(char* str)
{
	char*	pStr = new char[strlen(str)+1];
	strcpy(pStr,str);
	return pStr;
}
