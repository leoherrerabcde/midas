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


void pulse_export::Sheet_File_Constructor(void)
{
	m_Sheet_File.Celda	= NULL;
	m_Sheet_File.Cols	= 0;
	m_Sheet_File.Rows	= 0;
	m_Sheet_File.pFile	= NULL;
}

void pulse_export::Sheet_File_Destructor(void)
{
	if (m_Sheet_File.Celda!=NULL)
	{
		delete []m_Sheet_File.Celda	;
	}
	if (m_Sheet_File.pFile!=NULL)
	{
		fclose(m_Sheet_File.pFile);
		m_Sheet_File.pFile	= NULL;
	}
}

void pulse_export::Set_Sheet_Name	(char* XlsName,long SheetNumber)
{
	char	FileName[250];
	strcpy(FileName,XlsName);
	FileName[strlen(FileName)-4]	= '\0';
	sprintf(m_Sheet_File.SheetName,"%s_Sheet_%03d.bin",FileName,SheetNumber);
}

bool pulse_export::write_header_to_worksheet(void)
{
	m_Sheet_File.pFile	= fopen(m_Sheet_File.SheetName,"wb");
	if (m_Sheet_File.pFile==NULL)
	{
		return false;
	} 
	return write_header_to_worksheet(m_Sheet_File.pFile);
}

void pulse_export::Set_Sheet_Parameters(__File_Pwd_St *p_st_File_Pwd,long SheetNumber)
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
	if (bNew==true)
	{
		if (m_Sheet_File.Celda!=NULL)
		{
			delete [] m_Sheet_File.Celda;
		}
		m_Sheet_File.Celda		= new __Celda[mColumnCount*m_Sheet_File.Rows];
	}
}

bool pulse_export::write_header_to_worksheet(FILE* pFile)
{
	char			strHeader[32];
	pulse_format	class_pulse_format;
	long			col, length;
	long			o_col;
	
	if (pFile==NULL)
	{
		return false;
	}
	
	fwrite(&(m_Sheet_File.Cols),1,sizeof(long),pFile);
	fwrite(&(m_Sheet_File.Rows),1,sizeof(long),pFile);

	for(col=0;col<PWD_FIELD_COUNT;col++)
	{
		if (mColumnEnable[col])
		{
			o_col = mColumnOrder[col];
			class_pulse_format.format_pwd_header(strHeader,col);
			length	= strlen(strHeader);
			fwrite(&col,1,sizeof(col),pFile);		
			fwrite(&o_col,1,sizeof(o_col),pFile);		
			fwrite(&length,1,sizeof(length),pFile);		
			fwrite(strHeader,length,sizeof(char),pFile);		
		}
	}
	return true;
}

bool pulse_export::write_pwd_to_worksheet(__File_Pwd_St *p_st_File_Pwd)
{
	unsigned long	row;

	m_Correlative		= p_st_File_Pwd->p_st_Pwd_NewFields->ul_Index;
	m_dRelToa_ms		= p_st_File_Pwd->p_st_Pwd_NewFields->d_Rel_Toa_ms;

	for (row=0;row<p_st_File_Pwd->l_Pulse_Count;row++)
	{
		write_row_to_worksheet(p_st_File_Pwd,row);
		m_PlsCount ++;
		if (m_BreakXlsProcess==true)
		{
			break;
		}
	}
	
	return true;
}

void pulse_export::write_row_to_worksheet(__File_Pwd_St *p_st_File_Pwd, 
										  unsigned long	row)
{
	pulse_format	class_pulse_format;
	unsigned short	col,o_col;
	__Pwd_St		*p_st_Pwd;
	__Pwd_NF_St		*p_st_PwdNF;
	__Celda*		p_Celda;
	__Celda*		p_Col;
	
	p_st_Pwd	= p_st_File_Pwd->p_st_Pwd + row;
	p_st_PwdNF	= p_st_File_Pwd->p_st_Pwd_NewFields + row;
	p_Celda		= m_Sheet_File.Celda+row*m_Sheet_File.Cols;

	for(col=0;col<PWD_FIELD_COUNT;col++)
	{
		if (mColumnEnable[col])
		{
			o_col = mColumnOrder[col];
			//cell = sheet->Cell(row+1, o_col);
			p_Col	= p_Celda + o_col;
			switch(col){
			case 0:
				//cell->SetInteger(p_st_Pwd->uc_Adjust);
				p_Col->lValue	= p_st_Pwd->uc_Adjust;
				break;
			case 1:
				//cell->SetInteger(p_st_Pwd->uc_State );
				p_Col->lValue	= p_st_Pwd->uc_State;
				break;
			case 3:
				//cell->SetInteger(p_st_Pwd->us_Aoa);
				p_Col->lValue	= p_st_Pwd->us_Aoa;
				break;
			case 4:
				//cell->SetInteger(p_st_Pwd->us_Synth);
				p_Col->lValue	= p_st_Pwd->us_Synth;
				break;
			case 7:
				//cell->SetDouble(class_pulse_format.format_pwd_field(p_st_File_Pwd ,row,col));
				p_Col->dValue	= class_pulse_format.format_pwd_field(p_st_File_Pwd ,row,col);
				break;
			case 8:
				//cell->SetInteger(p_st_Pwd->ul_ToaCorregido);
				p_Col->lValue	= p_st_Pwd->ul_ToaCorregido;
				break;
			case COL_UL_REL_INDEX:
				//cell->SetInteger(p_st_PwdNF->ul_Rel_Index - m_Correlative);
				p_Col->lValue	= p_st_PwdNF->ul_Rel_Index - m_Correlative + 1;
				break;
			case COL_D_REL_S_TOA_MS:
				p_Col->dValue	= p_st_PwdNF->d_Rel_S_Toa_ms - m_dRelToa_ms;
				break;
			case 9:
			case 14:
			case 15:
				//fmt_gral.set_format_string(XLS_FORMAT_DATETIME);
				//cell->SetDouble(class_pulse_format.format_pwd_field(p_st_File_Pwd ,row,col));
				p_Col->dValue	= class_pulse_format.format_pwd_field(p_st_File_Pwd ,row,col);
				break;
			default:
				//cell->SetDouble(class_pulse_format.format_pwd_field(p_st_File_Pwd ,row,col));
				p_Col->dValue	= class_pulse_format.format_pwd_field(p_st_File_Pwd ,row,col);
			}
		}
	}
}

bool pulse_export::Save_Sheet(FILE* pFile)
{
	if (pFile==NULL)
	{
		return false;
	}
	
	fwrite(m_Sheet_File.Celda,m_Sheet_File.Cols*m_Sheet_File.Rows,sizeof(__Celda),pFile);
	fclose(pFile);
	return true;
}
