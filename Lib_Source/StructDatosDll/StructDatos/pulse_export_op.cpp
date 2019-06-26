/*
 * pulse_export_op.cpp
 *
 *  Created on: Sep 13, 2013
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

bool pulse_export::export_spread_sheet_filtered_op(char * p_ch_FullSpreadSheetFilename, 
												__File_Pwd_List_St *p_st_File_Pwd_List)
{
	BasicExcelWorksheet		*sheet;
	__File_Pwd_St			*p_st_FilePwd;
	bool					lvBreakDesable = false;
	bool*					bNewSheet;
	int						iNumWorkSheets;
	int						i;
	DWORD					lvTickEnd;
	char					lvStr[4096];
	DWORD					lvTickIni		= GetTickCount();
	
	m_PlsCount				= 0;
	m_Save_Count			= 0;
	m_Saving				= false;
	m_Creating				= true;
	strcpy(m_Xls_FileName,p_ch_FullSpreadSheetFilename);
	m_p_st_FilePwdList		= p_st_File_Pwd_List;
	
	p_st_FilePwd			= p_st_File_Pwd_List->p_st_FilePwdArray;
	bNewSheet				= new bool[p_st_File_Pwd_List->us_ListCount];
	for (i=0;i<p_st_File_Pwd_List->us_ListCount;i++)
	{
		bNewSheet[i]		= false;
	}
	iNumWorkSheets = m_xls.GetTotalWorkSheets();
	if (iNumWorkSheets==0)
	{
		m_xls.New(p_st_File_Pwd_List->us_ListCount);
		iNumWorkSheets		= p_st_File_Pwd_List->us_ListCount;
		for (i=0;i<iNumWorkSheets;i++)
		{
			sheet			= m_xls.GetWorksheet(i);
			write_header_to_worksheet_filtered(sheet);
			bNewSheet[i]	= true;
		}
	} 
	for (i=iNumWorkSheets;i<p_st_File_Pwd_List->us_ListCount;i++)
	{
		sheet				= m_xls.AddWorksheet(i);
		write_header_to_worksheet_filtered(sheet);
		bNewSheet[i]		= true;
	}
	for (i=iNumWorkSheets;i>p_st_File_Pwd_List->us_ListCount;i--)
	{
		m_xls.DeleteWorksheet(i-1);
	}
	m_xls.mSaveStatus			= &m_Save_Count;
	
	XLSFormatManager fmt_mgr(m_xls);
	CellFormat fmt_gral(fmt_mgr);
	fmt_gral.set_format_string(XLS_FORMAT_DATETIME);
	m_fmt_date_time				= &fmt_gral;
	
	m_IndexFile	= 0;
	for(i=0;i<p_st_File_Pwd_List->us_ListCount;i++)
	{
		m_PlsCount	= 0;
		sheet	= m_xls.GetWorksheet(i);
		Set_Sheet_Parameters_Op((p_st_FilePwd+i),i,sheet);
		//write_pwd_to_worksheet_filtered_op((p_st_FilePwd+i),sheet,&fmt_mgr,bNewSheet[i]);
		write_pwd_to_worksheet_filtered_op((p_st_FilePwd+i),sheet,bNewSheet[i]);
		lvTickEnd	= GetTickCount();
		sprintf(lvStr,"WorkSheet(%d)\t%d\t%d\t%d\n",
			i,
			lvTickIni,
			lvTickEnd,
			lvTickEnd-lvTickIni);
		(*m_pFnLog)(lvStr);
		lvTickIni	= lvTickEnd;
		m_IndexFile ++;
		if (m_BreakXlsProcess==true)
		{
			m_Save_Done = true;
			return false;
		}
	}
	delete [] bNewSheet;
	m_Creating	= false;
	m_Save_Done	= false;
	long SheetCount			= m_xls.GetTotalWorkSheets();
	m_xls.mBreakSave		= &lvBreakDesable;
	m_xls.SaveAs(m_Xls_FileName);
	m_xls.Close();
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

void pulse_export::Set_Sheet_Parameters_Op(__File_Pwd_St *p_st_File_Pwd,
										   long SheetNumber,
										   BasicExcelWorksheet* sheet)
{
	bool	bNew	= false;
	
	if(m_Sheet_File.Rows		!= p_st_File_Pwd->l_Pulse_Count)
	{
		m_Sheet_File.Rows		= p_st_File_Pwd->l_Pulse_Count;
		bNew		= true;
	}
	if (m_Sheet_File.Cols		!= mColumnCount)
	{
		bNew		= true;
		m_Sheet_File.Cols		= mColumnCount;
	}
	int SheetTotalRows	= sheet->GetTotalRows();
	BasicExcelCell*		cell;
	for (int i = SheetTotalRows ;i>m_Sheet_File.Rows;i--)
	{
		for(int j=0;j<sheet->GetTotalCols();j++)
		{
			cell		= sheet->Cell(i,j);
			cell->EraseContents();
		}
	}
}

bool pulse_export::write_pwd_to_worksheet_filtered_op(__File_Pwd_St *p_st_File_Pwd, 
												   BasicExcelWorksheet *sheet,
												   XLSFormatManager* pFmt,
												   bool bSheetNew)
{
	unsigned long	row;
	
	m_Correlative		= p_st_File_Pwd->p_st_Pwd_NewFields->ul_Rel_Index;
	m_dRelToa_ms		= p_st_File_Pwd->p_st_Pwd_NewFields->d_Rel_S_Toa_ms;

	for (row=0;row<p_st_File_Pwd->l_Pulse_Count;row++)
	{
		write_row_to_worksheet_filtered_op(p_st_File_Pwd,sheet,row,pFmt,bSheetNew);
		m_PlsCount ++;
		if (m_BreakXlsProcess==true)
		{
			break;
		}
	}
	
	return true;
}

bool pulse_export::write_pwd_to_worksheet_filtered_op(__File_Pwd_St *p_st_File_Pwd, 
													  BasicExcelWorksheet *sheet,
													  bool bSheetNew)
{
	unsigned long	row;
	
	m_Correlative		= p_st_File_Pwd->p_st_Pwd_NewFields->ul_Rel_Index;
	m_dRelToa_ms		= p_st_File_Pwd->p_st_Pwd_NewFields->d_Rel_S_Toa_ms;
	
	for (row=0;row<p_st_File_Pwd->l_Pulse_Count;row++)
	{
		write_row_to_worksheet_filtered_op(p_st_File_Pwd,sheet,row,bSheetNew);
		m_PlsCount ++;
		if (m_BreakXlsProcess==true)
		{
			break;
		}
	}
	
	return true;
}

void pulse_export::write_row_to_worksheet_filtered_op(__File_Pwd_St *p_st_File_Pwd, 
												   BasicExcelWorksheet *sheet, 
												   unsigned long	row,
												   XLSFormatManager* pFmt,
												   bool bSheetNew)
{
	pulse_format	class_pulse_format;
	unsigned short	col,o_col;
	__Pwd_St		*p_st_Pwd;
	__Pwd_NF_St		*p_st_PwdNF;
	BasicExcelCell	*cell;
	
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
				if (bSheetNew==true)
				{
					CellFormat fmt_gral(*pFmt);
					fmt_gral.set_format_string(XLS_FORMAT_DATETIME);
					cell->SetFormat(fmt_gral);
				}
				cell->SetDouble(class_pulse_format.format_pwd_field(p_st_File_Pwd ,row,col));
				break;
				break;
			default:
				cell->SetDouble(class_pulse_format.format_pwd_field(p_st_File_Pwd ,row,col));
			}
		}
	}
}

void pulse_export::write_row_to_worksheet_filtered_op(__File_Pwd_St *p_st_File_Pwd, 
													  BasicExcelWorksheet *sheet, 
													  unsigned long	row,
													  bool bSheetNew)
{
	pulse_format	class_pulse_format;
	unsigned short	col,o_col;
	__Pwd_St		*p_st_Pwd;
	__Pwd_NF_St		*p_st_PwdNF;
	BasicExcelCell	*cell;
	
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
				if (bSheetNew==true)
				{
					cell->SetFormat(*m_fmt_date_time);
				}
				cell->SetDouble(class_pulse_format.format_pwd_field(p_st_File_Pwd ,row,col));
				break;
				break;
			default:
				cell->SetDouble(class_pulse_format.format_pwd_field(p_st_File_Pwd ,row,col));
			}
		}
	}
}

