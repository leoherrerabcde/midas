/*
 * pulseexport.cpp
 *
 *  Created on: Sep 19, 2012
 *      Author: lherrera
 */

#include "pulse_conv_struct_define.h"
#include "filelist.h"
#include "pulseexport.h"
#include "pulseformat.h"


#include <stdio.h>
#include <stdlib.h>
#include <time.h>
#include <string.h>

typedef	void (CALLBACK* DLLFNVOID)(void);
typedef	void (CALLBACK* DLLFNGRAL)(int,char**);
typedef	void (CALLBACK* DLLFNCHAR)(char*);
typedef	void (CALLBACK* DLLFNLONGCHAR)(long,char*);
typedef	void (CALLBACK* DLLFNARRAY)(int,long*);
typedef	void (CALLBACK* DLLFNSHEET)(int,char*,__Sheet_File*);

extern DLLFNVOID			m_DllFnConstructor;
extern DLLFNVOID			m_DllFnDestructor;
extern DLLFNGRAL			m_DllFnSetHeader;
extern DLLFNSHEET			m_DllFnSetSheet;
extern DLLFNCHAR			m_DllFnSaveBook;
extern DLLFNARRAY			m_DllFnSetOrder;
extern DLLFNLONGCHAR		m_DllFnCvtBin;

char*		m_Header[PWD_FIELD_COUNT];
long		m_ArrayOrder[PWD_FIELD_COUNT];

void pulse_export::export_spread_sheet_ext(char * p_ch_SpreadSheetFileName, 
										   __File_Pwd_List_St *p_st_File_Pwd_List)
{
	long					i;
	__File_Pwd_St			*p_st_FilePwd;
	
	m_DllFnConstructor();
	write_new_book_ext();
	for (i=0;i<PWD_FIELD_COUNT;i++)
	{
		m_Header[i]	= new char[100];
	}

	m_PlsCount				= 0;
	m_Save_Count			= 0;
	m_Saving				= false;
	m_Creating				= true;
	m_p_st_FilePwdList		= p_st_File_Pwd_List;
	
	p_st_FilePwd			= p_st_File_Pwd_List->p_st_FilePwdArray;
	
	m_IndexFile	= 0;
	write_header_to_worksheet_ext();
	for(i=0;i<p_st_File_Pwd_List->us_ListCount;i++)
	{
		m_PlsCount	= 0;
		Set_Sheet_Parameters((p_st_FilePwd+i),i);
		write_pwd_to_worksheet((p_st_FilePwd+i));
		m_DllFnSetSheet(i,NULL,&m_Sheet_File);
		m_IndexFile ++;
	}
	
	m_DllFnSaveBook(p_ch_SpreadSheetFileName);
	m_DllFnDestructor();
	for (i=0;i<PWD_FIELD_COUNT;i++)
	{
		delete [] m_Header[i];
	}
	m_Save_Done = true;
}

void pulse_export::write_new_book_ext(void)
{
}

void pulse_export::write_name_worksheet_ext(void)
{
	
}

void pulse_export::write_header_to_worksheet_ext(void)
{
	pulse_format	class_pulse_format;
	long			col;
	long			o_col;
	long			Count;
	
	Count		= 0;
	for(col=0;col<PWD_FIELD_COUNT;col++)
	{
		if (mColumnEnable[col])
		{
			o_col	= mColumnOrder[col];
			m_ArrayOrder[o_col]	= col;
			class_pulse_format.format_pwd_header(m_Header[o_col],col);
			Count	++;
		}
	}
	m_DllFnSetOrder(Count,mColumnOrder);
	m_DllFnSetHeader(Count,m_Header);
}

void pulse_export::write_pwd_to_worksheet_ext(__File_Pwd_St *p_st_File_Pwd)
{
	
}

void pulse_export::write_row_to_worksheet_ext(__File_Pwd_St *p_st_File_Pwd,unsigned long row)
{
	
}

void pulse_export::save_book_ext(void)
{

}
