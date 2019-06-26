/*
 * pulseformat.cpp
 *
 *  Created on: Sep 19, 2012
 *      Author: lherrera
 */

#include "pulseformat.h"
#include <stdio.h>
#include <stdlib.h>
#include <time.h>
#include <string.h>

pulse_format::pulse_format() {
	// TODO Auto-generated constructor stub

}

pulse_format::~pulse_format() {
	// TODO Auto-generated destructor stub
}

double * pulse_format::format_pwd(__File_Pwd_St *p_st_File_Pwd, 
								  unsigned long Index, double * PwdArray)
{
	short i;

	for(i=0;i<PWD_FIELD_COUNT;i++)
	{
		PwdArray[i]		= format_pwd_field(p_st_File_Pwd,Index,i);
	}
	return PwdArray;
}

char * pulse_format::format_pwd(char *str_src,__File_Pwd_St *p_st_File_Pwd, unsigned long Index)
{
	char	strField[128];
	short	i;

	for(i=0;i<PWD_FIELD_COUNT;i++)
	{
		format_pwd_field(strField,p_st_File_Pwd,Index,i);
		if(i)
		{
			strcat(str_src,LIST_SEPARATE_STR);
			strcat(str_src,strField);
		}
		else
		{
			strcpy(str_src,strField);
		}
	}
	return str_src;
}

char * pulse_format::format_pwd_field(char *str_src,__File_Pwd_St *p_st_File_Pwd, unsigned long Index, short Field)
{
	if (Field<9)
	{
		__Pwd_St	*p_st_Pwd	= p_st_File_Pwd->p_st_Pwd + Index;
		switch (Field)
		{
		case 0:
			sprintf(str_src,"%d",p_st_Pwd->uc_Adjust);
			break;
		case 1:
			sprintf(str_src,"%d",p_st_Pwd->uc_State);
			break;
		case 2:
			sprintf(str_src,"%.3f",(float)p_st_Pwd->si_Amplitud/100.0);
			break;
		case 3:
			sprintf(str_src,"%d",p_st_Pwd->us_Aoa);
			break;
		case 4:
			sprintf(str_src,"%d",p_st_Pwd->us_Synth);
			break;
		case 5:
			sprintf(str_src,"%.2f",(float)p_st_Pwd->us_PulseWidth/100.0);
			break;
		case 6:
			sprintf(str_src,"%.1f",(float)p_st_Pwd->ul_IFM/10.0);
			break;
		case 7:
			sprintf(str_src,"%lu",p_st_Pwd->ul_Toa);
			break;
		case 8:
			sprintf(str_src,"%lu",p_st_Pwd->ul_ToaCorregido);
			break;
		}
	}
	else
	{
		__Pwd_NF_St	*p_st_PwdNF	= p_st_File_Pwd->p_st_Pwd_NewFields + Index;
		switch (Field)
		{
		case 9:
			//format_d_Date(str_src,p_st_PwdNF->d_Date_s);
			format_abs_time(str_src,p_st_PwdNF->d_Date_s);
			break;
		case 10:
			if(Index==3412){
				Index=Index;
			}
			sprintf(str_src,"%.2f",p_st_PwdNF->d_pre__d_Toa_us);
			break;
		case 11:
			sprintf(str_src,"%.2f",p_st_PwdNF->d_post_d_Toa_us);
			break;
		case 12:
			sprintf(str_src,"%.3f",p_st_PwdNF->d_Rel_Toa_ms);
			break;
		case 13:
			sprintf(str_src,"%.3f",p_st_PwdNF->d_Rel_S_Toa_ms);
			break;
		case 14:
			format_abs_time(str_src,p_st_PwdNF->d_Abs_Toa_s);
			break;
		case 15:
			format_abs_time(str_src,p_st_PwdNF->d_FileTime_s);
			break;
		case 16:
			sprintf(str_src,"%d",p_st_PwdNF->us_FileNumber);
			break;
		case 17:
			sprintf(str_src,"%d",p_st_PwdNF->uc_RollOver);
			break;
		case 18:
			sprintf(str_src,"%d",p_st_PwdNF->uc_PulseDetail);
			break;
		case 19:
			sprintf(str_src,"%d",p_st_PwdNF->uc_WrapAround);
			break;
		case 20:
			sprintf(str_src,"%d",p_st_PwdNF->uc_Toa_Error);
			break;
		case 21:
			sprintf(str_src,"%d",p_st_PwdNF->st_ProcessError.ErrorCode);
			break;
		case PWD_FIELD_COUNT-2:
			sprintf(str_src,"%d",p_st_PwdNF->ul_Rel_Index);
			break;
		case PWD_FIELD_COUNT-1:
			sprintf(str_src,"%d",p_st_PwdNF->ul_Index);
			break;
		}
	}
	return str_src;
}

double pulse_format::format_pwd_field(__File_Pwd_St *p_st_File_Pwd, unsigned long Index, short Field)
{
	if (p_st_File_Pwd->l_Pulse_Count<=0)
	{
		Field=Field;
		return -1;
	}
	if (Field<9)
	{
		__Pwd_St	*p_st_Pwd	= p_st_File_Pwd->p_st_Pwd + Index;
		switch (Field)
		{
		case 0:
			return (double)(p_st_Pwd->uc_Adjust);
			break;
		case 1:
			return (double)(p_st_Pwd->uc_State);
			break;
		case 2:
			return (double)p_st_Pwd->si_Amplitud/100.0;
			break;
		case 3:
			return (double)(p_st_Pwd->us_Aoa);
			break;
		case 4:
			return (double)(p_st_Pwd->us_Synth);
			break;
		case 5:
			return (double)p_st_Pwd->us_PulseWidth/100.0;
			break;
		case 6:
			return (double)p_st_Pwd->ul_IFM/10.0;
			break;
		case 7:
			if (CONVERT_TOA)
			{
				return ((double)p_st_Pwd->ul_Toa)/100.0;
			} 
			else
			{
				return (double)(p_st_Pwd->ul_Toa);
			}
			break;
		case 8:
			return (double)(p_st_Pwd->ul_ToaCorregido);
			break;
		}
	}
	else
	{
		__Pwd_NF_St	*p_st_PwdNF	= p_st_File_Pwd->p_st_Pwd_NewFields + Index;
		switch (Field)
		{
		case 9:
			//return p_st_PwdNF->d_Date_s;
			return (p_st_PwdNF->d_Date_s+OFFSET_TIME)/ABS_TOA_TO_DATE;
			break;
		case 10:
			return (double)(p_st_PwdNF->d_pre__d_Toa_us);
			break;
		case 11:
			return (double)(p_st_PwdNF->d_post_d_Toa_us);
			break;
		case 12:
			return (double)(p_st_PwdNF->d_Rel_Toa_ms);
			break;
		case 13:
			return (double)(p_st_PwdNF->d_Rel_S_Toa_ms);
			break;
		case 14:
			return (p_st_PwdNF->d_Abs_Toa_s+OFFSET_TIME)/ABS_TOA_TO_DATE;
			break;
		case 15:
			return (p_st_PwdNF->d_FileTime_s+OFFSET_TIME)/ABS_TOA_TO_DATE;
			break;
		case 16:
			return (double)(p_st_PwdNF->us_FileNumber);
			break;
		case 17:
			return (double)(p_st_PwdNF->uc_RollOver);
			break;
		case 18:
			return (double)(p_st_PwdNF->uc_PulseDetail);
			break;
		case 19:
			return (double)(p_st_PwdNF->uc_WrapAround);
			break;
		case 20:
			return (double)(p_st_PwdNF->uc_Toa_Error);
			break;
		case 21:
			return (double)(p_st_PwdNF->st_ProcessError.ErrorCode);
			break;
		case PWD_FIELD_COUNT-2:
			return (double)(p_st_PwdNF->ul_Rel_Index);
			break;
		case PWD_FIELD_COUNT-1:
			return (double)(p_st_PwdNF->ul_Index);
			break;
		}
	}
	return 0.0;
}

char * pulse_format::format_d_Date(char *str_src,double d_Date)
{
	time_t		lTime	= (time_t)d_Date;
	struct tm 	*p_st_tm;

	if (lTime < 0)
	{
		return NULL;
	}
	p_st_tm	= localtime(&lTime);
	sprintf(str_src,"%04d/%02d/%02d %02d:%02d:%02d",p_st_tm->tm_year+1900,p_st_tm->tm_mon,p_st_tm->tm_mday,p_st_tm->tm_hour,p_st_tm->tm_min,p_st_tm->tm_sec);
	return str_src;
}

char * pulse_format::format_pwd_header(char *str_dst)
{
	char	strHeader[128];
	short	i;

	for(i=0;i<PWD_FIELD_COUNT;i++)
	{
		format_pwd_header(strHeader,i);
		if(i)
		{
			strcat(str_dst,LIST_SEPARATE_STR);
			strcat(str_dst,strHeader);
		}
		else
		{
			strcpy(str_dst,strHeader);
		}
	}
	return str_dst;
}

char * pulse_format::format_pwd_header(char *str_dst, short IndexHeader)
{
	if (IndexHeader<9)
	{
		switch (IndexHeader)
		{
		case 0:
			strcpy(str_dst,"Adjust");
			break;
		case 1:
			strcpy(str_dst,"State");
			break;
		case 2:
			strcpy(str_dst,"Amp[dBm]");
			break;
		case 3:
			strcpy(str_dst,"AOA[°]");
			break;
		case 4:
			strcpy(str_dst,"Synth[MHz]");
			break;
		case 5:
			strcpy(str_dst,"PW[usec]");
			break;
		case 6:
			strcpy(str_dst,"IFM[MHz]");
			break;
		case 7:
			strcpy(str_dst,"TOA");
			break;
		case 8:
			strcpy(str_dst,"TOA_CORREGIDO");
			break;
		}
	}
	else
	{
		switch (IndexHeader)
		{
		case 9:
			strcpy(str_dst,"Date");
			break;
		case 10:
			strcpy(str_dst,"pre_DToa[us]");
			break;
		case 11:
			strcpy(str_dst,"post_DToa[us]");
			break;
		case 12:
			strcpy(str_dst,"Rel_Toa[ms]");
			break;
		case 13:
			strcpy(str_dst,"Sheet_Rel_Toa[ms]");
			break;
		case 14:
			strcpy(str_dst,"Abs_Toa[s]");
			break;
		case 15:
			strcpy(str_dst,"FileTime[s]");
			break;
		case 16:
			strcpy(str_dst,"FileNumber");
			break;
		case 17:
			strcpy(str_dst,"RollOverEvent");
			break;
		case 18:
			strcpy(str_dst,"PulseDetailEvent");
			break;
		case 19:
			strcpy(str_dst,"WranpAroundEvent");
			break;
		case 20:
			strcpy(str_dst,"Toa_Error_Code");
			break;
		case 21:
			strcpy(str_dst,"ErrorCode");
			break;
		case PWD_FIELD_COUNT-2:
			strcpy(str_dst,"Relative");
			break;
		case PWD_FIELD_COUNT-1:
			strcpy(str_dst,"Correlative");
			break;
		}
	}
	return str_dst;
}

char * pulse_format::format_abs_time(char * dest, double d_abs_time)
{
	sprintf(dest,"%f",(double)((d_abs_time+OFFSET_TIME)/ABS_TOA_TO_DATE));
	return dest;
}
