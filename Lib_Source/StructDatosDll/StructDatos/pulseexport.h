/*
 * pulseexport.h
 *
 *  Created on: Sep 19, 2012
 *      Author: lherrera
 */

#ifndef PULSEEXPORT_H_
#define PULSEEXPORT_H_

#include "pulse_conv_struct_define.h"
#include "ExcelFormat.h"
#include "Pulseproject.h"
// #include "DllFnType.h"

using namespace ExcelFormat;

class pulse_export {
public:
	pulse_export();
	virtual ~pulse_export();

	void 				set_spread_sheet ();
	
	bool	export_spread_sheet_filtered(char * p_ch_SpreadSheetPath, char *p_ch_SpreadSheetName, __File_Pwd_List_St *p_st_File_Pwd_List);
	bool				export_spread_sheet(char * p_ch_SpreadSheetPath, char *p_ch_SpreadSheetName, __File_Pwd_List_St *p_st_File_Pwd_List);
	bool				export_spread_sheet(char * p_ch_SpreadSheetFileName, __File_Pwd_List_St *p_st_File_Pwd_List);
	bool				export_file(__File_Pwd_List_St *p_st_File_Pwd_List);
	bool				export_file(char * p_ch_PathExport, __File_Pwd_List_St *p_st_File_Pwd_List);
	bool				export_file(char * p_ch_FileName_Export, __File_Pwd_St *p_st_File_Pwd);
	void				save_xls_file(void);
	void				Set_SpreadSheet_File(char *FileName){strcpy(m_Xls_FileName,FileName);}
	void	Run_CreateXls(void);
	void	Run_CreateXlsOp(void);
	void	Run_CreateXls(bool bGenBin);
	void	Run_CreateXls(Pulse_project* pPulseProject);
	void	Set_FilePwdList(__File_Pwd_List_St *p){m_p_st_FilePwdList=p;}
	long	Get_PlsCount(void) {return m_PlsCount;}
	long	Get_IndexFile(void) {
		long	lvIndexFile = (long)m_IndexFile;
		return lvIndexFile;}
	void	Clr_IndexFile(){m_IndexFile=0;}
	long	Get_IndexSpreadFile(void) {return (long)m_IndexSpreadFile;}
	bool	Get_SSCreation_Done(){return m_Create_Done;}
	bool	Get_SS_Save_Done(){return m_Save_Done;}
	void	Clr_SSCreation_Done(){m_Create_Done=false;}
	void	Clr_SS_Save_Done(){m_Save_Done=false;}
	void	Set_SS_Save_Done(){m_Save_Done=true;}
	void	setColumnFormat		(long* ColumnOrder, long* ColumnEnable);
	bool	export_spread_sheet_filtered(char * p_ch_SpreadSheetFileName, 
										__File_Pwd_List_St *p_st_File_Pwd_List);
	bool	export_spread_sheet_filtered(char * p_ch_SpreadSheetFileName, 
										__File_Pwd_List_St *p_st_File_Pwd_List,
										BasicExcel* xls);
	bool	CancelXlsProcess(void);

	void	SetBinGenState(bool bState);
	bool	GetBinGenState(void);

	/*void Set_DllConstructor((DLLFNVOID pFn){ m_DllFnConstructor=pFn;};
	void Set_DllDestructor(DLLFNVOID pFn){ m_DllFnDestructor=pFn;};*/
	void	(*m_pFnLog)(char*);
	DWORD					m_TickCount;

private:

	bool	export_file					(char * p_ch_FileName, 
										 char * p_ch_PathName, 
										 __File_Pwd_St *p_st_File_Pwd );

	bool	write_header_to_worksheet	(void);
	bool	write_header_to_worksheet(FILE* pFile);
	bool	write_header_to_worksheet	(BasicExcelWorksheet *sheet);
	bool	write_header_to_worksheet_filtered(BasicExcelWorksheet *sheet);

	bool	write_pwd_to_worksheet		(__File_Pwd_St *p_st_File_Pwd);
	bool	write_pwd_to_worksheet		(__File_Pwd_St *p_st_File_Pwd, 
										 BasicExcelWorksheet *sheet);
	bool	write_pwd_to_worksheet		(__File_Pwd_St *p_st_File_Pwd, 
										 BasicExcelWorksheet *sheet,
										 XLSFormatManager* pFmt);
	bool write_pwd_to_worksheet_filtered (__File_Pwd_St *p_st_File_Pwd, 
										 BasicExcelWorksheet *sheet,
										 XLSFormatManager* pFmt);

	void	write_row_to_worksheet		(__File_Pwd_St *p_st_File_Pwd,unsigned long row);
	void	write_row_to_worksheet		(__File_Pwd_St *p_st_File_Pwd, 
										 BasicExcelWorksheet *sheet ,
										 unsigned long	row);
	void	write_row_to_worksheet		(__File_Pwd_St *p_st_File_Pwd, 
										 BasicExcelWorksheet *sheet ,
										 unsigned long	row,
										 XLSFormatManager* pFmt);
	void write_row_to_worksheet_filtered(__File_Pwd_St *p_st_File_Pwd, 
										BasicExcelWorksheet *sheet ,
										unsigned long	row,
										XLSFormatManager* pFmt);
	void Set_Sheet_Name	(char* XlsName,long SheetNumber);
	void Set_Sheet_Parameters(__File_Pwd_St *p_st_File_Pwd,long SheetNumber);
	bool Save_Sheet(FILE* pFile);
	void Sheet_File_Constructor(void);
	void Sheet_File_Destructor(void);

	void export_spread_sheet_ext		(char * p_ch_SpreadSheetFileName, 
										 __File_Pwd_List_St *p_st_File_Pwd_List);
	void write_new_book_ext				(void);
	void write_name_worksheet_ext		(void);
	void write_header_to_worksheet_ext	(void);
	void write_pwd_to_worksheet_ext		(__File_Pwd_St *p_st_File_Pwd);
	void write_row_to_worksheet_ext		(__File_Pwd_St *p_st_File_Pwd,unsigned long row);
	void save_book_ext					(void);

	bool export_spread_sheet_filtered_op(char * p_ch_FullSpreadSheetFilename, 
										__File_Pwd_List_St *p_st_File_Pwd_List);
	void Set_Sheet_Parameters_Op		(__File_Pwd_St *p_st_File_Pwd,long SheetNumber,BasicExcelWorksheet* sheet);
	bool write_pwd_to_worksheet_filtered_op	(__File_Pwd_St *p_st_File_Pwd, 
											BasicExcelWorksheet *sheet,
											XLSFormatManager* pFmt,
											bool SheetNew);
	bool write_pwd_to_worksheet_filtered_op	(__File_Pwd_St *p_st_File_Pwd, 
											BasicExcelWorksheet *sheet,
											bool SheetNew);
	void write_row_to_worksheet_filtered_op(__File_Pwd_St *p_st_File_Pwd, 
											BasicExcelWorksheet *sheet, 
											unsigned long	row,
											XLSFormatManager* pFmt,
											bool SheetNew);
	void write_row_to_worksheet_filtered_op(__File_Pwd_St *p_st_File_Pwd, 
											BasicExcelWorksheet *sheet, 
											unsigned long	row,
											bool SheetNew);

	//void _Destroy						

	BasicExcel				m_xls;
	CellFormat*				m_fmt_date_time;
	unsigned long			m_PlsCount;
	long					m_Save_Count;
	unsigned_long			m_IndexFile;
	unsigned_long			m_IndexSpreadFile;
	bool					m_Saving;
	bool					m_Creating;
	char					m_Xls_FileName[260];
	__File_Pwd_List_St		*m_p_st_FilePwdList;
	bool					m_Create_Done;
	bool					m_Save_Done;
	long					mColumnOrder[PWD_FIELD_COUNT];
	long					mColumnEnable[PWD_FIELD_COUNT];
	long					mColumnCount;
	bool					m_BreakXlsProcess;
	bool					m_ProcessCanceled;
	__Sheet_File			m_Sheet_File;
	bool					m_BinGenEnable;

	unsigned_long			m_Correlative;
	double					m_dRelToa_ms;

	/*DLLFNVOID				m_DllFnConstructor;
	DLLFNVOID				m_DllFnDestructor;*/
};

#endif /* PULSEEXPORT_H_ */
