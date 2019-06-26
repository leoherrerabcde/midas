/*
 * pulseexportexcel.cpp
 *
 *  Created on: Sep 21, 2012
 *      Author: lherrera
 */

#include "pulse_conv_struct_define.h"
#include "filelist.h"
#include "pulseformat.h"
#include <stdio.h>
#include <stdlib.h>
#include <time.h>
#include <string.h>
#include "pulseexportexcel.h"
#include "mexcel.h"
using namespace miniexcel;

pulse_export_excel::pulse_export_excel() {
	// TODO Auto-generated constructor stub
	strcpy(m_ch_ExcelExtension,".xls");
}

pulse_export_excel::~pulse_export_excel() {
	// TODO Auto-generated destructor stub
}

// Private Functions
bool pulse_export_excel::export_spread_sheet(char * p_ch_SpreadSheetName, char * p_ch_SpreadSheetPath, __File_Pwd_St *p_st_File_Pwd )
{
	char	ch_FullFileName[256];

	strcpy(ch_FullFileName,p_ch_SpreadSheetPath);
	strcat(ch_FullFileName,"\\");
	strcat(ch_FullFileName,p_ch_SpreadSheetName);
	export_spread_sheet(ch_FullFileName,p_st_File_Pwd);
	return true;
}

// Public Functions
bool pulse_export_excel::export_spread_sheet(__File_Pwd_List_St *p_st_File_Pwd_List)
{
	return export_spread_sheet(p_st_File_Pwd_List->p_PathName, p_st_File_Pwd_List);
}

bool pulse_export_excel::export_spread_sheet(char * p_ch_SpreadSheetPath, __File_Pwd_List_St *p_st_File_Pwd_List)
{
	file_list			class_FileList;
	__File_Pwd_St		*p_st_FilePwd;
	char				strFileExport[256];
	unsigned short		i;

	p_st_FilePwd				= p_st_File_Pwd_List->p_st_FilePwdArray;

	for(i=0;i<p_st_File_Pwd_List->us_ListCount;i++)
	{
		strcpy(strFileExport,p_st_FilePwd->p_ch_FileName);
		class_FileList.change_file_extension(strFileExport,m_ch_ExcelExtension);
		export_spread_sheet(strFileExport,p_ch_SpreadSheetPath,p_st_FilePwd++);
	}
	return true;
}

bool pulse_export_excel::export_spread_sheet(char * p_ch_SpreadSheetName, __File_Pwd_St *p_st_File_Pwd)
{
	FILE 			*pFile;
	char			strData[1024];
	pulse_format	class_pulse_format;
	CMiniExcel 		classExcel;
	unsigned long	ul_Row,i,j;

	pFile = fopen(p_ch_SpreadSheetName,"wb");
	if (pFile == NULL)
	{
		return false;
	}
	strcat(strData,"\n");
	fwrite(strData,sizeof(char),strlen(strData),pFile);

	for(i=0;i<p_st_File_Pwd->l_Pulse_Count;i++)
	{
		class_pulse_format.format_pwd(strData,p_st_File_Pwd,i);
		strcat(strData,"\n");
		fwrite(strData,sizeof(char),strlen(strData),pFile);
	}
	fclose(pFile);
	return true;
}
