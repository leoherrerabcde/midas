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


bool pulse_export::export_spread_sheet_filtered(char * p_ch_FullSpreadSheetFilename, 
												__File_Pwd_List_St *p_st_File_Pwd_List)
{
	BasicExcel				xls;
	short					i;
	BasicExcelWorksheet		*sheet;
	__File_Pwd_St			*p_st_FilePwd;
	bool					lvBreakDesable = false;
	DWORD					lvTickEnd;
	char					lvStr[4096];

	DWORD		lvTickIni	= GetTickCount();
	m_PlsCount				= 0;
	m_Save_Count			= 0;
	m_Saving				= false;
	m_Creating				= true;
	strcpy(m_Xls_FileName,p_ch_FullSpreadSheetFilename);
	m_p_st_FilePwdList		= p_st_File_Pwd_List;
	
	p_st_FilePwd			= p_st_File_Pwd_List->p_st_FilePwdArray;
	xls.New(p_st_File_Pwd_List->us_ListCount);
	xls.mSaveStatus			= &m_Save_Count;

	XLSFormatManager fmt_mgr(xls);
	m_IndexFile	= 0;
	for(i=0;i<p_st_File_Pwd_List->us_ListCount;i++)
	{
		m_PlsCount	= 0;
		sheet	= xls.GetWorksheet(i);
		write_header_to_worksheet_filtered(sheet);
		write_pwd_to_worksheet_filtered((p_st_FilePwd+i), sheet,&fmt_mgr);
		m_IndexFile ++;
		if (m_BreakXlsProcess==true)
		{
			m_Save_Done = true;
			return false;
		}
		lvTickEnd	= GetTickCount();
		sprintf(lvStr,"WorkSheet(%d)\t%d\t%d\t%d\n",
			i,
			lvTickIni,
			lvTickEnd,
			lvTickEnd-lvTickIni);
		(*m_pFnLog)(lvStr);
		lvTickIni	= lvTickEnd;
	}
	
	m_Creating	= false;
	m_Save_Done	= false;
	long SheetCount	= xls.GetTotalWorkSheets();
	//xls.mBreakSave	= &m_BreakXlsProcess;
	xls.mBreakSave		= &lvBreakDesable;
	xls.SaveAs(m_Xls_FileName);
	xls.Close();
	lvTickEnd	= GetTickCount();
	sprintf(lvStr,"Saving Time\t%d\t%d\t%d\n",
		lvTickIni,
		lvTickEnd,
		lvTickEnd-lvTickIni);
	(*m_pFnLog)(lvStr);
	lvTickIni	= lvTickEnd;
	m_TickCount	= lvTickEnd;
	long Index=1;
	m_Save_Done = true;
	return true;
}


/*bool pulse_export::export_project()
{

}*/


bool pulse_export::export_spread_sheet_filtered(char * p_ch_FullSpreadSheetFilename, 
												__File_Pwd_List_St *p_st_File_Pwd_List,
												BasicExcel* xls)
{
	short					i;
	BasicExcelWorksheet		*sheet;
	__File_Pwd_St			*p_st_FilePwd;
	
	m_PlsCount				= 0;
	m_Save_Count			= 0;
	m_Saving				= false;
	m_Creating				= true;
	strcpy(m_Xls_FileName,p_ch_FullSpreadSheetFilename);
	m_p_st_FilePwdList		= p_st_File_Pwd_List;
	
	p_st_FilePwd			= p_st_File_Pwd_List->p_st_FilePwdArray;
	xls->New(p_st_File_Pwd_List->us_ListCount);
	xls->mSaveStatus			= &m_Save_Count;
	
	XLSFormatManager fmt_mgr(*xls);
	m_IndexFile	= 0;
	for(i=0;i<p_st_File_Pwd_List->us_ListCount;i++)
	{
		m_PlsCount	= 0;
		sheet	= xls->GetWorksheet(i);
		write_header_to_worksheet_filtered(sheet);
		write_pwd_to_worksheet_filtered((p_st_FilePwd+i), sheet,&fmt_mgr);
		m_IndexFile ++;
	}
	
	m_Creating	= false;
	m_Save_Done	= false;
	long SheetCount	= xls->GetTotalWorkSheets();
	xls->SaveAs(m_Xls_FileName);
	xls->Close();
	long Index=1;
	m_Save_Done = true;
	return true;
}


bool pulse_export::write_header_to_worksheet_filtered(BasicExcelWorksheet *sheet)
{
	char			strHeader[32];
	pulse_format	class_pulse_format;
	int				col, row = 0;
	int				o_col;
	BasicExcelCell* cell;
	
	row = 0;
	/*sprintf(strHeader,"NumPls");
	cell = sheet->Cell(row, 0);		
	cell->Set(strHeader);*/
	for(col=0;col<PWD_FIELD_COUNT;col++)
	{
		if (mColumnEnable[col])
		{
			o_col = mColumnOrder[col];
			class_pulse_format.format_pwd_header(strHeader,col);
			cell = sheet->Cell(row, o_col);		
			cell->Set(strHeader);
		}
	}
	return false;
}


bool pulse_export::write_pwd_to_worksheet_filtered(__File_Pwd_St *p_st_File_Pwd, 
										  BasicExcelWorksheet *sheet,
										  XLSFormatManager* pFmt)
{
	// pulse_format	class_pulse_format;// 
	unsigned long	row;

	m_Correlative		= p_st_File_Pwd->p_st_Pwd_NewFields->ul_Rel_Index;
	m_dRelToa_ms		= p_st_File_Pwd->p_st_Pwd_NewFields->d_Rel_Toa_ms;
	for (row=0;row<p_st_File_Pwd->l_Pulse_Count;row++)
	{
		write_row_to_worksheet_filtered(p_st_File_Pwd,sheet,row,pFmt);
		m_PlsCount ++;
		if (m_BreakXlsProcess==true)
		{
			break;
		}
	}
	
	return true;
}


void pulse_export::write_row_to_worksheet_filtered(__File_Pwd_St *p_st_File_Pwd, 
										  BasicExcelWorksheet *sheet, 
										  unsigned long	row,
										  XLSFormatManager* pFmt)
{
	pulse_format	class_pulse_format;
	unsigned short	col,o_col;
	__Pwd_St		*p_st_Pwd;
	__Pwd_NF_St		*p_st_PwdNF;
	BasicExcelCell	*cell;
	
	CellFormat fmt_gral(*pFmt);
	p_st_Pwd	= p_st_File_Pwd->p_st_Pwd + row;
	p_st_PwdNF	= p_st_File_Pwd->p_st_Pwd_NewFields + row;
	for(col=0;col<PWD_FIELD_COUNT;col++)
	{
		if (mColumnEnable[col])
		{
			o_col = mColumnOrder[col];
			cell = sheet->Cell(row+1, o_col);
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
			case COL_UL_REL_INDEX:
				cell->SetInteger(p_st_PwdNF->ul_Rel_Index - m_Correlative);
				break;
			case COL_D_REL_S_TOA_MS:
				cell->SetDouble(p_st_PwdNF->d_Rel_S_Toa_ms - m_dRelToa_ms);
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
}

bool pulse_export::CancelXlsProcess(void)
{
	if(m_Save_Done==true && m_Create_Done==true)
	{
		return true;
	}
	m_BreakXlsProcess	= true;
	return m_ProcessCanceled;
}