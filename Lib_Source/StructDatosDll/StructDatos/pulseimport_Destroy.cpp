/*
 * pulseimport.cpp
 *
 *  Created on: oct 23, 2012
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

void pulse_import::_FieldDestroy(__File_Pwd_List_St* pFilePwdLst)
{
	if(pFilePwdLst!=NULL)
	{
		if(pFilePwdLst->p_PathName)
		{
			_Destroy(pFilePwdLst->p_PathName);
		}
		if(pFilePwdLst->p_st_FileList)
		{
			_Destroy(pFilePwdLst->p_st_FileList);
		}
		if (pFilePwdLst!=NULL)
		{
			_DestroyArraySt(pFilePwdLst->p_st_FilePwdArray,pFilePwdLst->us_ListCount);
			// _DestroyArray(pFilePwdLst->p_st_FilePwdArray);
		}
	}
}

void pulse_import::_DestroyArray(__File_Pwd_St* pFilePwdArray)
{
// 	if (pFilePwdArray)
// 	{
// 		delete	[]pFilePwdArray;
// 	}
	m_PoolFilePwd->Free(pFilePwdArray);
}

void pulse_import::_DestroyArraySt(__File_Pwd_St* pFilePwdArray,long lvCount)
{
	long			i;
	__File_Pwd_St*	pIt;

	if (pFilePwdArray)
	{
		pIt	= pFilePwdArray;
		for (i=0;i<lvCount;i++)
		{
			_DestroyArray(pIt->p_ch_FileName);
			_DestroyArray(pIt->p_st_Pwd);
			_DestroyArray(pIt->p_st_Pwd_NewFields);
			pIt->p_ch_FileName		= NULL;
			pIt->p_st_Pwd			= NULL;
			pIt->p_st_Pwd_NewFields	= NULL;
			pIt++;
		}
// 		delete	[]pFilePwdArray;
		m_PoolFilePwd->Free(pFilePwdArray);
	}
}
