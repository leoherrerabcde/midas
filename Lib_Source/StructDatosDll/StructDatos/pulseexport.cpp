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
#include "ExcelFormat.h"

using namespace ExcelFormat;

#include <stdio.h>
#include <stdlib.h>
#include <time.h>
#include <string.h>

pulse_export::pulse_export() {
	// TODO Auto-generated constructor stub
	m_PlsCount		= 0;
	m_Save_Count	= 0;
	m_Saving		= false;
	m_Creating		= false;
	m_IndexFile		= 0;
	sprintf(m_Xls_FileName,"");
	m_p_st_FilePwdList	= NULL;
	m_ProcessCanceled	= false;
	m_BreakXlsProcess	= false;
	m_BinGenEnable		= false;
	Sheet_File_Constructor();
}

pulse_export::~pulse_export() {
	// TODO Auto-generated destructor stub
}

void pulse_export::set_spread_sheet ()
{

}

// Private Functions
bool pulse_export::export_file(char * p_ch_FileName, char * p_ch_PathName, __File_Pwd_St *p_st_File_Pwd )
{
	char	ch_FullFileName[256];

	strcpy(ch_FullFileName,p_ch_PathName);
	strcat(ch_FullFileName,"\\");
	strcat(ch_FullFileName,p_ch_FileName);
	export_file(ch_FullFileName,p_st_File_Pwd);
	return true;
}

// Public Functions
bool pulse_export::export_file(__File_Pwd_List_St *p_st_File_Pwd_List)
{
	return export_file(p_st_File_Pwd_List->p_PathName, p_st_File_Pwd_List);
}

bool pulse_export::export_file(char * p_ch_PathExport, __File_Pwd_List_St *p_st_File_Pwd_List)
{
	file_list			class_FileList;
	__File_Pwd_St		*p_st_FilePwd;
	char				strFileExport[256];
	char				strCVS[5];
	unsigned short		i;

	p_st_FilePwd				= p_st_File_Pwd_List->p_st_FilePwdArray;
	strcpy(strCVS,".csv");

	for(i=0;i<p_st_File_Pwd_List->us_ListCount;i++)
	{
		strcpy(strFileExport,p_st_FilePwd->p_ch_FileName);
		class_FileList.change_file_extension(strFileExport,strCVS);
		export_file(strFileExport,p_ch_PathExport,p_st_FilePwd++);
	}
	return true;
}

bool pulse_export::export_file(char * p_ch_FileName_Export, __File_Pwd_St *p_st_File_Pwd)
{
	FILE 			*pFile;
	long	i;
	char			strData[4096];
	pulse_format	class_pulse_format;
	long			lDebug=3412;

	pFile = fopen(p_ch_FileName_Export,"wb");
	if (pFile == NULL)
	{
		return false;
	}
	class_pulse_format.format_pwd_header(strData);
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

bool pulse_export::export_spread_sheet(char * p_ch_SpreadSheetPath, char *p_ch_SpreadSheetName, __File_Pwd_List_St *p_st_File_Pwd_List)
{
	char		FullExcelFileName[260];

	sprintf(FullExcelFileName,"%s\\%s",p_ch_SpreadSheetPath,p_ch_SpreadSheetName);
	return export_spread_sheet(FullExcelFileName, p_st_File_Pwd_List);
}

void pulse_export::write_row_to_worksheet(__File_Pwd_St *p_st_File_Pwd, 
										  BasicExcelWorksheet *sheet, 
										  unsigned long	row)
{
	pulse_format	class_pulse_format;
	unsigned short	col;
	__Pwd_St		*p_st_Pwd;
	BasicExcelCell	*cell;
//	static long		bugCount=0;

	p_st_Pwd	= p_st_File_Pwd->p_st_Pwd + row;
	for(col=0;col<PWD_FIELD_COUNT;col++)
	{
		cell = sheet->Cell(row+1, col+1);
		if (p_st_File_Pwd->l_Pulse_Count<=0)
		{
			col=col;
			break;
		}
		switch(col){
			case 0:
				cell->SetInteger(p_st_Pwd->uc_Adjust);
				break;
			case 1:
				cell->SetInteger(p_st_Pwd->uc_State );
				break;
			case 3:
				cell->SetInteger(p_st_Pwd->us_Aoa);
				break;
			case 4:
				cell->SetInteger(p_st_Pwd->us_Synth);
				break;
			case 7:
				if (CONVERT_TOA)
				{
					cell->SetDouble(class_pulse_format.format_pwd_field(p_st_File_Pwd ,row,col));
				} 
				else
				{
					cell->SetInteger(p_st_Pwd->ul_Toa);
				}
				break;
			case 8:
				cell->SetInteger(p_st_Pwd->ul_ToaCorregido);
				break;
			default:
				cell->SetDouble(class_pulse_format.format_pwd_field(p_st_File_Pwd ,row,col));
		}
	}
}

void pulse_export::write_row_to_worksheet(__File_Pwd_St *p_st_File_Pwd, 
										  BasicExcelWorksheet *sheet, 
										  unsigned long	row,
										  XLSFormatManager* pFmt)
{
	pulse_format	class_pulse_format;
	unsigned short	col;
	__Pwd_St		*p_st_Pwd;
	BasicExcelCell	*cell;
	//	static long		bugCount=0;
	
	CellFormat fmt_gral(*pFmt);
	p_st_Pwd	= p_st_File_Pwd->p_st_Pwd + row;
	for(col=0;col<PWD_FIELD_COUNT;col++)
	{
		if (p_st_File_Pwd->l_Pulse_Count<=0)
		{
			col=col;
			break;
		}
		cell = sheet->Cell(row+1, col+1);
		switch(col){
		case 0:
			cell->SetInteger(p_st_Pwd->uc_Adjust);
			break;
		case 1:
			cell->SetInteger(p_st_Pwd->uc_State );
			break;
		case 3:
			cell->SetInteger(p_st_Pwd->us_Aoa);
			break;
		case 4:
			cell->SetInteger(p_st_Pwd->us_Synth);
			break;
		case 7:
			cell->SetDouble(class_pulse_format.format_pwd_field(p_st_File_Pwd ,row,col));
			//cell->SetFormat(fmt_gral);
			break;
		case 8:
			cell->SetInteger(p_st_Pwd->ul_ToaCorregido);
			break;
		case 9:
		case 14:
		case 15:
			fmt_gral.set_format_string(XLS_FORMAT_DATETIME);
			cell->SetDouble(class_pulse_format.format_pwd_field(p_st_File_Pwd ,row,col));
			cell->SetFormat(fmt_gral);
			break;
			break;
		default:
			cell->SetDouble(class_pulse_format.format_pwd_field(p_st_File_Pwd ,row,col));
		}
	}
}

bool pulse_export::write_pwd_to_worksheet(__File_Pwd_St *p_st_File_Pwd, BasicExcelWorksheet *sheet)
{
	// pulse_format	class_pulse_format;// 
	unsigned long	row;

	for (row=0;row<p_st_File_Pwd->l_Pulse_Count;row++)
	{
		if (p_st_File_Pwd->l_Pulse_Count==0)
		{
			p_st_File_Pwd->l_Pulse_Count=p_st_File_Pwd->l_Pulse_Count;
		}
		write_row_to_worksheet(p_st_File_Pwd,sheet,row);
		m_PlsCount ++;
	}
	
	return true;
}

bool pulse_export::write_pwd_to_worksheet(__File_Pwd_St *p_st_File_Pwd, 
										  BasicExcelWorksheet *sheet,
										  XLSFormatManager* pFmt)
{
	// pulse_format	class_pulse_format;// 
	unsigned long	row;
	
	for (row=0;row<p_st_File_Pwd->l_Pulse_Count;row++)
	{
		if (p_st_File_Pwd->l_Pulse_Count==0)
		{
			p_st_File_Pwd->l_Pulse_Count=p_st_File_Pwd->l_Pulse_Count;
		}
		write_row_to_worksheet(p_st_File_Pwd,sheet,row,pFmt);
		m_PlsCount ++;
	}
	
	return true;
}

bool pulse_export::write_header_to_worksheet(BasicExcelWorksheet *sheet)
{
	char			strHeader[32];
	pulse_format	class_pulse_format;
	int				col, row = 0;
	BasicExcelCell* cell;

	row = 0;
	sprintf(strHeader,"NumPls");
	cell = sheet->Cell(row, 0);		
	cell->Set(strHeader);
	for(col=0;col<PWD_FIELD_COUNT;col++)
	{
		class_pulse_format.format_pwd_header(strHeader,col);
		cell = sheet->Cell(row, col+1);		
		cell->Set(strHeader);
	}
	return false;
}


bool pulse_export::export_spread_sheet(char * p_ch_FullSpreadSheetFilename, 
									   __File_Pwd_List_St *p_st_File_Pwd_List)
{
	long					i;
	__File_Pwd_St			*p_st_FilePwd;
	
	m_PlsCount				= 0;
	m_Save_Count			= 0;
	m_Saving				= false;
	m_Creating				= true;
	m_p_st_FilePwdList		= p_st_File_Pwd_List;
	
	p_st_FilePwd			= p_st_File_Pwd_List->p_st_FilePwdArray;

	m_IndexFile	= 0;
	for(i=0;i<p_st_File_Pwd_List->us_ListCount;i++)
	{
		m_PlsCount	= 0;
		Set_Sheet_Name(p_ch_FullSpreadSheetFilename,i);
		Set_Sheet_Parameters((p_st_FilePwd+i),i);
		write_header_to_worksheet();
		write_pwd_to_worksheet((p_st_FilePwd+i));
		Save_Sheet(m_Sheet_File.pFile);
		m_Sheet_File.pFile	= NULL;
		m_IndexFile ++;
	}
	
	long Index=1;
	Sheet_File_Destructor();
	m_Save_Done = true;
	return true;
}


void pulse_export::save_xls_file(void)
{
	m_Save_Done	= false;
	m_xls.SaveAs(m_Xls_FileName);
	m_Save_Done = true;
}

void pulse_export::Run_CreateXls(bool bGenBin)
{
	m_Create_Done	= false;
	//export_spread_sheet_ext(m_Xls_FileName,m_p_st_FilePwdList);
	if (bGenBin==true)
	{
		export_spread_sheet(m_Xls_FileName,m_p_st_FilePwdList);
	} 
	else
	{
		export_spread_sheet_filtered(m_Xls_FileName,m_p_st_FilePwdList);
	}
	m_Create_Done	= true;
	if (m_BreakXlsProcess==true)
	{
		m_ProcessCanceled	= true;
	}
	else
	{
		m_Create_Done	= true;
	}
}

void pulse_export::Run_CreateXls(void)
{
	m_Create_Done	= false;
	//export_spread_sheet_ext(m_Xls_FileName,m_p_st_FilePwdList);
	if (m_BinGenEnable==true)
	{
		export_spread_sheet(m_Xls_FileName,m_p_st_FilePwdList);
	} 
	else
	{
		export_spread_sheet_filtered(m_Xls_FileName,m_p_st_FilePwdList);
	}
	m_Create_Done	= true;
	if (m_BreakXlsProcess==true)
	{
		m_ProcessCanceled	= true;
	}
	else
	{
		m_Create_Done	= true;
	}
}


void pulse_export::Run_CreateXlsOp(void)
{
	m_Create_Done	= false;
	export_spread_sheet_filtered_op(m_Xls_FileName,m_p_st_FilePwdList);
	m_Create_Done	= true;
	if (m_BreakXlsProcess==true)
	{
		m_ProcessCanceled	= true;
	}
	else
	{
		m_Create_Done	= true;
	}
}

void pulse_export::Run_CreateXls(Pulse_project* pPulseProject)
{
	BasicExcel				xls;
	char					OutputFileName[260];
	char*					pFileName;
	char					FileNumber[10];
	__File_Pwd_List_St*		p_st_FilePwdList = NULL;

	m_Create_Done	= false;
	strcpy(OutputFileName,pPulseProject->mProject.outputPath->FileName);
	strcat(OutputFileName,"\\");
	strcat(OutputFileName,pPulseProject->mProject.missionName->FileName);
	strcat(OutputFileName,"_00000.xls");
	pFileName		= &OutputFileName[strlen(OutputFileName)-9];

	for (m_IndexSpreadFile=0;
		 m_IndexSpreadFile<pPulseProject->GetSpreadFileCount();
		 m_IndexSpreadFile++)
	{
		sprintf(FileNumber,"%05d",m_IndexSpreadFile);	
		memcpy(pFileName,FileNumber,5);
		if (!pPulseProject->Create_SpreadFile(p_st_FilePwdList,m_IndexSpreadFile))
		{
			p_st_FilePwdList	= pPulseProject->Create_SpreadFile(m_IndexSpreadFile);
		}
		Set_SpreadSheet_File(OutputFileName);
		Clr_SSCreation_Done();
		Clr_SS_Save_Done();
		Set_FilePwdList(p_st_FilePwdList);
	}
	export_spread_sheet_filtered(m_Xls_FileName,m_p_st_FilePwdList,&xls);
	m_Create_Done	= true;
}


void pulse_export::setColumnFormat(long* ColumnOrder, long* ColumnEnable)
{
	memcpy(mColumnOrder,ColumnOrder,PWD_FIELD_COUNT*sizeof(long));
	memcpy(mColumnEnable,ColumnEnable,PWD_FIELD_COUNT*sizeof(long));
	mColumnCount		= 0;
	
	for (long i=0;i<PWD_FIELD_COUNT;i++)
	{
		if (mColumnEnable[i])
		{
			mColumnCount	++;
		}
	}
}

void pulse_export::SetBinGenState(bool bState)
{
	if(bState==false)
	{
		m_BinGenEnable	= false;
	}
	else
	{
		m_BinGenEnable	= true;
	}
	//m_BinGenEnable		= bState;
}

bool pulse_export::GetBinGenState(void)
{
	return m_BinGenEnable;
}