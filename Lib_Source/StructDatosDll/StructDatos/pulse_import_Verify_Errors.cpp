/*
 * pulseimport.cpp
 *
 *  Created on: Sep 16, 2012
 *      Author: lherrera
 */

#include "pulse_conv_struct_define.h"
#include "pulseimport.h"
#include "filelist.h"

#include <stdio.h>
#include <stdlib.h>
#include <time.h>
#include <dirent.h>
#include <string.h>
#include <iostream>
#include <list>
using namespace std;

unsigned_long	LV_Count = 0;

unsigned char pulse_import::VerifyPwdErrors(__Pwd_St* p_st_Pwd,
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
		/*__PwdPointerSt	st_ErrPrt;

		st_ErrPrt.d_Abs_Toa_s		= p_st_PwdNF->d_Abs_Toa_s;
		st_ErrPrt.d_post_d_Toa_us	= p_st_PwdNF->d_post_d_Toa_us;
		st_ErrPrt.d_Rel_Toa_ms		= p_st_PwdNF->d_Rel_Toa_ms;
		st_ErrPrt.s_Error_Code		= sCodeError.ErrorCode;
		st_ErrPrt.ul_Index			= p_st_PwdNF->ul_Index;
		st_ErrPrt.ul_Toa			= p_st_Pwd->ul_Toa;
		st_ErrPrt.ul_IndexFile		= p_st_PwdNF->us_FileNumber;
		st_ErrPrt.ul_IndexPwd		= ul_IndexPwd;

		m_ErrorPtrLst.push_back(st_ErrPrt);*/
		if (m_ErrPrtCnt)
		{
			if (m_ErrPrtCnt<m_ErrPrtMaxCnt)
			{
				m_pErrPrt++;
			} 
			else
			{
				return sCodeError.ErrorCode;
			}
		} 
		else
		{
			m_pErrPrt = m_ErrorPtrLstPt;
		}
		m_pErrPrt->d_Abs_Toa_s		= p_st_PwdNF->d_Abs_Toa_s;
		m_pErrPrt->d_post_d_Toa_us	= p_st_PwdNF->d_post_d_Toa_us;
		m_pErrPrt->d_Rel_Toa_ms		= p_st_PwdNF->d_Rel_Toa_ms;
		m_pErrPrt->s_Error_Code		= sCodeError.ErrorCode;
		m_pErrPrt->ul_Index			= p_st_PwdNF->ul_Index;
		m_pErrPrt->ul_Toa			= p_st_Pwd->ul_Toa;
		m_pErrPrt->ul_IndexFile		= p_st_PwdNF->us_FileNumber;
		m_pErrPrt->ul_IndexPwd		= ul_IndexPwd;
		//m_pErrPrt->ul_IndexFilePwdList	= 
		//m_pErrPrt					++;
		m_ErrPrtCnt					++;
		LV_Count					++;
	}
	return sCodeError.ErrorCode;
}

short pulse_import::VerifyNewDate(double NewDate, double PreviusDate)
{
	if (NewDate<PreviusDate)
	{
		return 1;
	}
	return 0;
}

short pulse_import::VerifyFileDate(double NewDate_seg, double FileDate_seg)
{
	double DiffDate_seg = NewDate_seg - FileDate_seg;

	if ((DiffDate_seg<-MAX_DIFF_FILE_TIME_SEC) || (DiffDate_seg>MAX_DIFF_FILE_TIME_SEC))
	{
		return 1;
	}
	return 0;
}


short pulse_import::VerifyDToaError(double dDToa_us)
{
	if (dDToa_us<=0)
	{
		return 1;
	}
	return 0;
}

short pulse_import::VerifyRToaError(double d_RToa,double d_RToa_Previus)
{
	if (d_RToa<=d_RToa_Previus)
	{
		return 1;
	}
	return 0;
}

short pulse_import::VerifyPW(unsigned short usPW)
{
	return VerifyRange(usPW, PWD_PW_MIN, PWD_PW_MAX);
}

short pulse_import::VerifyAmplitud(short sAmp)
{
	return VerifyRange(sAmp, PWD_AMP_MIN, PWD_AMP_MAX);
}

short pulse_import::VerifyFrec(unsigned_long lFrec)
{
	return VerifyRange(lFrec, PWD_FREC_MIN, PWD_FREC_MAX);
}

short pulse_import::VerifyRange(unsigned_long ulVal, unsigned_long ulMin, unsigned_long ulMax)
{
	if ((ulVal < ulMin) || (ulVal > ulMax))
	{
		return 1;
	}
	return 0;
}

short pulse_import::VerifyRange(unsigned short ulVal, unsigned short ulMin, unsigned short ulMax)
{
	if ((ulVal < ulMin) || (ulVal > ulMax))
	{
		return 1;
	}
	return 0;
}

short pulse_import::VerifyRange(long lVal, long lMin, long lMax)
{
	if ((lVal < lMin) || (lVal > lMax))
	{
		return 1;
	}
	return 0;
}

short pulse_import::VerifyRange(short lVal, short lMin, short lMax)
{
	if ((lVal < lMin) || (lVal > lMax))
	{
		return 1;
	}
	return 0;
}

