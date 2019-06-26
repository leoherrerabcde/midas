
#include <windows.h>
#include <string.h>
#include <stdio.h>
#include <stdlib.h>
#include <time.h>
#include <list>

#ifdef _MSC_VER
#include <crtdbg.h>
#endif

#include "pulse_conv_struct_define.h"
#include "gralFunctions.h"
#include "filelist.h"
#include "pulseexport.h"
#include "pulseformat.h"
#include "pulseimport.h"
#include "Pulseproject.h"
#include "PoolMemory.h"
#include "ErrorHandler.h"
// #include "DllFnType.h"

//#define		__BACKGROUND__
typedef	void (CALLBACK* DLLFNVOID)(void);
typedef	void (CALLBACK* DLLFNGRAL)(int,char**);
typedef	void (CALLBACK* DLLFNCHAR)(char*);
typedef	void (CALLBACK* DLLFNLONGCHAR)(long,char*);
typedef	void (CALLBACK* DLLFNSHEET)(int,char*,__Sheet_File*);
typedef	void (CALLBACK* DLLFNARRAY)(int,long*);

DLLFNVOID			m_DllFnConstructor;
DLLFNVOID			m_DllFnDestructor;
DLLFNCHAR			m_DllFnSaveBook;
DLLFNGRAL			m_DllFnSetHeader;
DLLFNSHEET			m_DllFnSetSheet;
DLLFNARRAY			m_DllFnSetOrder;
DLLFNLONGCHAR		m_DllFnCvtBin;

pulse_import		*m_pulse_import_class	= NULL;
__File_Pwd_List_St	*p_st_FilePwdList		= NULL;
pulse_export		*mPulseExport			= NULL;
Pulse_project		mPulseProject;
CErrorHandler		mErrorHandler;

HINSTANCE			m_hDll = NULL;
char				m_cvt2xls_dll[250];

CPoolMemory<__File_Pwd_List_St>		mPool_FilePwdLst;
CPoolMemory<__File_Pwd_St>			mPool_FilePwd;
CPoolMemory<__Pwd_NF_St>			mPool_PwdNF;
CPoolMemory<__Pwd_St>				mPool_Pwd;
CPoolMemory<__PwdPointerSt>			mPool_PwdErr;

FILE				*m_Hdl_File		= NULL;


UINT	ThreadCreateSpreadSheetOp( LPVOID pParam );
UINT	ThreadCreateSpreadSheet	( LPVOID pParam );
UINT	ThreadSaveSpreadSheet	( LPVOID pParam );
UINT	ThreadPulseImport		( LPVOID pParam );
UINT	ThreadCreateWorkSpace	( LPVOID pParam );
UINT	ThreadCreateSpreadProject( LPVOID pParam );
UINT	ThreadPulseInitPoolMemory( LPVOID pParam );

void	Pulse_Log_Write(char* lvData);
void	f_Log_Write(char *lsData);
void	f_Log_Write_Header(void);


void    ProccessMapFile(char* lsMapFile);
bool    IsMapFile(char* lsFileName);
int     ReadMapFile(char* lsMapFile, list<string> &lsFileList);
UINT	ThreadBinToXls( LPVOID pParam );
void	ConvertBin2Xls(string &strXls);
void	CvtBin2Xls(int SheetCnt,const char* strXls);

char*	GV_ErrorHeader[] = {"Error_Code",
							"Xls",
							"Sheet",
							"Pulse",
							"File",
							"Correlative",
							"Index",
							"Toa",
							"Post_DToa[us]",
							"Rel_Toa[ms]",
							"Abs_Toa"};

//extern void RunThread ( void *pfnThreadProc,LPVOID pParam);

// MyDll_ReverseString -- Reverses the characters of a given string
void __stdcall MyDll_ReverseString(LPSTR lpString)
{
	_strrev(lpString);
}


// MyDLL_Rotate -- Returns bit rotation of 32-bit integer value
int __stdcall MyDll_Rotate(int nVal, int nDirect, short iNumBits)
{
	int nRet = 0;
	
	if((iNumBits < 1) || (iNumBits > 31))
		return nRet;
	
	switch(nDirect)
	{
	case 0:
		// Rotate nVal left by iNumBits
		nRet = (((nVal) << (iNumBits)) |
			((nVal) >> (32-(iNumBits))));
		break;
	case 1:
		// Rotate nVal right by iNumBits
		nRet = (((nVal) >> (iNumBits)) |
			((nVal) << (32-(iNumBits))));
		break;
	}
	
	return nRet;
}


void Read_Data ( void )
{

}


int __stdcall DllFn_Compound_Number(unsigned char nByte0	,
									unsigned char nByte1	,
									unsigned char nByte2	,
									unsigned char nByte3	,
									unsigned char nDigits)
{
	int					lvValue	;
	unsigned char		* pVal;

	pVal				= (unsigned char *) ( & lvValue	);

	* pVal ++			=	nByte0	;

	if ( nDigits > 8 )
	{
		* pVal ++		=	nByte1	;
		if ( nDigits > 16 )
		{
			* pVal ++	=	nByte2	;
			if ( nDigits > 24 )
			{
				* pVal	=	nByte3	;
			}
			else
			{
				* pVal	=	0		;
			}
		}
		else
		{
			* pVal ++	=	0		;
			* pVal		=	0		;
		}
	}
	else
	{
		* pVal ++		=	0		;
		* pVal ++		=	0		;
		* pVal			=	0		;
	}
	return lvValue	;
}

int __stdcall DllFn_Compound_Number16(unsigned char nByte0	,
									unsigned char nByte1	)
{
	int					lvValue	;
	unsigned char		* pVal;
	
	pVal				= (unsigned char *) ( & lvValue	);
	
	* pVal ++			=	nByte0	;
	* pVal ++			=	nByte1	;
	* pVal ++			=	0		;
	* pVal				=	0		;
	return				lvValue	;
}

int __stdcall DllFn_Compound_Number24(unsigned char nByte0	,
									unsigned char nByte1	,
									unsigned char nByte2	)
{
	int					lvValue	;
	unsigned char		* pVal;
	
	pVal				= (unsigned char *) ( & lvValue	);
	
	* pVal ++			=	nByte0	;
	* pVal ++			=	nByte1	;
	* pVal ++			=	nByte2	;
	* pVal				=	0		;
	return				lvValue	;
}

int __stdcall DllFn_Compound_Number32(unsigned char nByte0	,
									unsigned char nByte1	,
									unsigned char nByte2	,
									unsigned char nByte3	)
{
	int					lvValue	;
	unsigned char		* pVal;
	
	pVal				= (unsigned char *) ( & lvValue	);
	
	* pVal ++			=	nByte0	;
	* pVal ++			=	nByte1	;
	* pVal ++			=	nByte2	;
	* pVal				=	nByte3	;
	return				lvValue	;
}

char	Conv_Digit_2_Hex	( unsigned char	lCh)
{
	if ( lCh < 10 )
	{
		return	(char) ( '0' + lCh )	;
	}
	else
	{
		if ( lCh	<	16 )
		{
			return	(char) ( 54 + lCh )	;
		} 
		else
		{
			return		'\0'	;
		}
	}
	return		'\0'	;
}

LPSTR __stdcall	DllFn_Conv_Dec_2_Hex ( int lDec , short lDigits )
{
	unsigned char		lCh			;
	unsigned char		i			;
	char				* pHex		;

	pHex				= new char[lDigits+1]	;
	
	pHex[lDigits]		= '\0'	;

	for ( i = 1		; i <= lDigits	; i++ )
	{
		lCh					=	lDec	&	0x0f	;
		pHex[lDigits-i]		=	Conv_Digit_2_Hex ( lCh );
		lDec				>>=	4;
	}
	
	return				pHex	;
}

void __stdcall	DllFn_Conv_Dec_2_Hex ( unsigned char	 * byteArray	, 
									   int lCount						, 
									   LPSTR pStrPrev	)
{
	//pStrPrev		=	new char [lCount]	;
	memcpy			( pStrPrev	,	byteArray	, lCount);
	//pStrOld			=	pStrPrev	;
	//return			pStrTmp			;
}

int __stdcall	DllFn_Conv_Str_2_IdEvent ( LPSTR	pStrEv , unsigned char	lDigits	)
{
	int				lIdEv	;
	//short			lVal	;
	short			* pIdEv	;
	unsigned char	i		;
	unsigned char	lBitRot	;
	unsigned char	lAsc	;

	pIdEv			= (short *)	& lIdEv	;
	lBitRot			=	0	;

	for ( i = 0 ; i < lDigits	; i ++ )
	{
		lAsc		= ( ( * pStrEv ++	) - 65 ) & 0x1f	;
		if ( i )
		{
			if ( lBitRot > 7 )
			{
				lBitRot	-=	8;
				* ++ pIdEv	|= ((short)( lAsc )) << ( lBitRot );
			} 
			else
			{
				* pIdEv		|=	((short)( lAsc )) << ( lBitRot );
			}
		}
		else
		{
			 * pIdEv 	=	( lAsc )	;
		}
		lBitRot			+=	5;
	}
	return				lIdEv	;
}

short __stdcall Pulse_Files_Count ( )
{
	if (m_pulse_import_class == NULL)
	{
		return -1;
	}
	return p_st_FilePwdList->us_ListCount;
}

long __stdcall Pulse_Count ( short IndexSheet )
{
	if (m_pulse_import_class == NULL)
	{
		return -1;
	}
	
	__File_Pwd_St	*p_st_FilePwd;
	
	if (p_st_FilePwdList->us_ListCount <= IndexSheet)
	{
		return -1;
	}
	p_st_FilePwd	= p_st_FilePwdList->p_st_FilePwdArray + IndexSheet;
	return p_st_FilePwd->l_Pulse_Count;
}

short __stdcall Pulse_Field_Count ( )
{
	return PWD_FIELD_COUNT;
}

void __stdcall Pulse_Field_Header ( short IndexField , LPSTR *StrFileName)
{
// 	if (m_pulse_import_class != NULL)
// 	{
// 		pulse_format	mPulseFormat;
// 		mPulseFormat.format_pwd_header(*StrFileName,IndexField);
// 	}
	pulse_format	mPulseFormat;
	mPulseFormat.format_pwd_header(*StrFileName,IndexField);
}
void __stdcall Pulse_Get_Pwd(short IndexSheet, long IndexPulse, double *Pwd)
{
	if (m_pulse_import_class != NULL)
	{
		pulse_format	mPulseFormat;
		__File_Pwd_St	*p_st_FilePwd;

		if ( IndexSheet < p_st_FilePwdList->us_ListCount)
		{
			p_st_FilePwd	= p_st_FilePwdList->p_st_FilePwdArray + IndexSheet;
			mPulseFormat.format_pwd(p_st_FilePwd,IndexPulse,Pwd);
		}
	}
}

void __stdcall Pulse_Get_File(long IndexSpread, long IndexSheet,LPSTR *StrFileName)
{
// 	if (m_pulse_import_class != NULL)
// 	{
// 		m_pulse_import_class->get_pulse_file(IndexFile,*StrFileName);
// 	}
	mPulseProject.get_FileName(IndexSpread,IndexSheet,*StrFileName);
}

void __stdcall Pulse_Finish_Xls(void)
{
	if (mPulseExport!=NULL)
	{
		delete mPulseExport;
		mPulseExport	= NULL;
	}
}

void __stdcall Pulse_Create_Xls_File_Op(LPSTR StrFileName,long IndexSpread)
{
	DWORD		lvTickIni	= GetTickCount();
	
	if (mPulseExport==NULL)
	{
		mPulseExport = new pulse_export;
		mPulseExport->m_pFnLog	= Pulse_Log_Write;
		f_Log_Write("\tTime\t\tTickIni\tTickEnd\tEnlased\n");
	}
	mPulseExport->SetBinGenState(false);
	
	mPulseProject.Destroy(p_st_FilePwdList);
	
	p_st_FilePwdList	= mPulseProject.Create_SpreadFile(IndexSpread);
	
	DWORD	lvTickEnd	= GetTickCount();
	char	lvStr[4096];
	sprintf(lvStr,"Create_SpreadFile\t%d\t%d\t%d\n",
		lvTickIni,
		lvTickEnd,
		lvTickEnd-lvTickIni);
	Pulse_Log_Write(lvStr);
	
	mPulseExport->Set_SpreadSheet_File(StrFileName);
	mPulseExport->Clr_SSCreation_Done();
	mPulseExport->Clr_SS_Save_Done();
	mPulseExport->Set_FilePwdList(p_st_FilePwdList);
	mPulseExport->Clr_IndexFile();
	mPulseExport->setColumnFormat(mPulseProject.mProject.workSheetConfiguration.ColumnOrder,
		mPulseProject.mProject.workSheetConfiguration.ColumnEnable);
	::CreateThread(NULL,0,(LPTHREAD_START_ROUTINE)ThreadCreateSpreadSheetOp,(LPVOID)(mPulseExport),0,NULL);
}

void __stdcall Pulse_Create_Xls_File(LPSTR StrFileName,long IndexSpread,bool bGenBinEnable)
{
	DWORD		lvTickIni	= GetTickCount();

	if (mPulseExport!=NULL)
	{
		//_Destroy(mPulseExport);
		delete mPulseExport;
		mPulseExport	= NULL;
	}
	else
	{
		f_Log_Write("\tTime\t\tTickIni\tTickEnd\tEnlased\n");
	}
	mPulseExport = new pulse_export;
	mPulseExport->SetBinGenState(bGenBinEnable);
	mPulseExport->m_pFnLog	= Pulse_Log_Write;

	mPulseProject.Destroy(p_st_FilePwdList);

	p_st_FilePwdList	= mPulseProject.Create_SpreadFile(IndexSpread);

	DWORD	lvTickEnd	= GetTickCount();
	char	lvStr[4096];
	sprintf(lvStr,"Create_SpreadFile\t%d\t%d\t%d\n",
			lvTickIni,
			lvTickEnd,
			lvTickEnd-lvTickIni);
	Pulse_Log_Write(lvStr);

	mPulseExport->Set_SpreadSheet_File(StrFileName);
	mPulseExport->Clr_SSCreation_Done();
	mPulseExport->Clr_SS_Save_Done();
	mPulseExport->Set_FilePwdList(p_st_FilePwdList);
	mPulseExport->setColumnFormat(mPulseProject.mProject.workSheetConfiguration.ColumnOrder,
								  mPulseProject.mProject.workSheetConfiguration.ColumnEnable);
	::CreateThread(NULL,0,(LPTHREAD_START_ROUTINE)ThreadCreateSpreadSheet,(LPVOID)(mPulseExport),0,NULL);
}


void __stdcall Pulse_Create_Xls_Project(void)
{
	if (mPulseExport!=NULL)
	{
		_Destroy(mPulseExport);
	}
	mPulseExport = new pulse_export;
	mPulseExport->setColumnFormat(mPulseProject.mProject.workSheetConfiguration.ColumnOrder,
							mPulseProject.mProject.workSheetConfiguration.ColumnEnable);
	::CreateThread(NULL,0,(LPTHREAD_START_ROUTINE)ThreadCreateSpreadProject,
				   (LPVOID)(&mPulseProject),0,NULL);
}

void __stdcall Pulse_Set_Xls_Dll(LPSTR StrFileName,
								 LPSTR sBookConstructor,
								 LPSTR sBookDestructor,
								 LPSTR sBookSave,
								 LPSTR sBookSetHeader,
								 LPSTR sBookSetSheet,
								 LPSTR sBookSetOrder,
								 LPSTR sBookCvtBin)
{
	strcpy(m_cvt2xls_dll,StrFileName);

	if (m_hDll==NULL)
	{
		m_hDll	= LoadLibrary(m_cvt2xls_dll);
		if (m_hDll!=NULL)
		{
			m_DllFnConstructor	= (DLLFNVOID)GetProcAddress(m_hDll,sBookConstructor);
			m_DllFnDestructor	= (DLLFNVOID)GetProcAddress(m_hDll,sBookDestructor);
			m_DllFnSaveBook		= (DLLFNCHAR)GetProcAddress(m_hDll,sBookSave);
			m_DllFnSetHeader	= (DLLFNGRAL)GetProcAddress(m_hDll,sBookSetHeader);
			m_DllFnSetSheet		= (DLLFNSHEET)GetProcAddress(m_hDll,sBookSetSheet);
			m_DllFnSetOrder		= (DLLFNARRAY)GetProcAddress(m_hDll,sBookSetOrder);
			m_DllFnCvtBin		= (DLLFNLONGCHAR)GetProcAddress(m_hDll,sBookCvtBin);
		} 
		else
		{
		}
	}
}

void __stdcall Pulse_Export_File(LPSTR StrFileName)
{
	if (m_pulse_import_class == NULL)
	{
		return;
	}
	if (mPulseExport==NULL)
	{
		pulse_export* p=new pulse_export;
		mPulseExport = p;
	}
	/*if (m_hDll==NULL)
	{
		m_hDll	= LoadLibrary(m_cvt2xls_dll);
	}*/
	mPulseExport->Set_SpreadSheet_File(StrFileName);
	mPulseExport->Clr_SSCreation_Done();
	mPulseExport->Clr_SS_Save_Done();
	mPulseExport->Set_FilePwdList(p_st_FilePwdList);

	/*mPulseExport->Set_DllConstructor(m_DllFnConstructor);
	mPulseExport->Set_DllDestructor(m_DllFnDestructor);*/

	::CreateThread(NULL,0,(LPTHREAD_START_ROUTINE)ThreadCreateSpreadSheet,(LPVOID)(mPulseExport),0,NULL);
	//mPulseExport->Run_CreateXls();
}

long __stdcall Pulse_Import_File_BG(LPSTR StrPath)
{
	pulse_import 		pulse_import_cmd;
	
	cFileName	pulse_path(StrPath);
	
	mPulseProject.Clear_ImportFileDone();
	mPulseProject.Set_FileCount(0);
	mPulseProject.setMissionPath(&(pulse_path.mName));

	::CreateThread(NULL,0,(LPTHREAD_START_ROUTINE)ThreadPulseImport,(LPVOID)(&mPulseProject),0,NULL);

	//char* pPath = mPulseProject.mProject.outputPath->FileName;
	
	/*mPool_FilePwdLst.DumpMemory(pPath,"mPool_FilePwdLst");
	mPool_FilePwd.DumpMemory(pPath,"mPool_FilePwd");
	mPool_PwdNF.DumpMemory(pPath,"mPool_PwdNF");
	mPool_Pwd.DumpMemory(pPath,"mPool_Pwd");*/

	/*char lsData[250];

	sprintf(lsData,"pIndexFilePulse = %d",mPulseProject.Get_Index_FilePulse());
	sprintf(lsData,"pIndexFilePwdList = %d",mPulseProject.Get_Index_FilePwdList());
	sprintf(lsData,"pFileCount = %d",mPulseProject.Get_FileCount());
	sprintf(lsData,"pDone = %d",mPulseProject.Get_ImportFileDone());*/

	return mPulseProject.mProject.FilePwdList_Count;
}


long __stdcall Pulse_Import_File(LPSTR StrPath)
{
	pulse_import 		pulse_import_cmd;
	
	cFileName	pulse_path(StrPath);

	//_CrtDumpMemoryLeaks();

	mPulseProject.setMissionPath(&(pulse_path.mName));
	//_CrtDumpMemoryLeaks();
	
	pulse_import_cmd.m_PoolFilePwdLst	= &mPool_FilePwdLst;
	pulse_import_cmd.m_PoolFilePwd		= &mPool_FilePwd;
	pulse_import_cmd.m_PoolPwdNF		= &mPool_PwdNF;
	pulse_import_cmd.m_PoolPwd			= &mPool_Pwd;
	pulse_import_cmd.m_PoolPwdErr		= &mPool_PwdErr;

	pulse_import_cmd.import_file(&mPulseProject);

	//char* pPath = mPulseProject.mProject.outputPath->FileName;

	return mPulseProject.mProject.FilePwdList_Count;
}

void __stdcall Pulse_Import_File_Status ( long* pIndexFilePwdList, 
										  long* pIndexFilePulse,
										  long* pFileCount,
										  long*	pDone)
{
	*pIndexFilePulse		= mPulseProject.Get_Index_FilePulse();
	*pIndexFilePwdList		= mPulseProject.Get_Index_FilePwdList();
	*pFileCount				= mPulseProject.Get_FileCount();
	*pDone					= mPulseProject.Get_ImportFileDone();
	char 	lsData[250];

	sprintf(lsData,"pIndexFilePulse = %d",*pIndexFilePulse);
	sprintf(lsData,"pIndexFilePwdList = %d",*pIndexFilePwdList);
	sprintf(lsData,"pFileCount = %d",*pFileCount);
	sprintf(lsData,"pDone = %d",*pDone);
}

long __stdcall Pulse_Sheets_Per_Pulses ( int PulseCount )
{
	mPulseProject.mProject.workSheetConfiguration.PulseQtyCriteria		= PulseCount;
	mPulseProject.mProject.workSheetConfiguration.IntervalTimeCriteria	= 0;
	//_CrtDumpMemoryLeaks();
	return 0;
}

void __stdcall Pulse_Sheets_Per_File ( long SheetCount )
{
	mPulseProject.mProject.workSheetConfiguration.workSheetsPerXlsCount	= SheetCount;
	//_CrtDumpMemoryLeaks();
}

long __stdcall Pulse_Sheets_Per_Interval ( double TimeInterval )
{
	mPulseProject.mProject.workSheetConfiguration.PulseQtyCriteria		= 0;
	mPulseProject.mProject.workSheetConfiguration.IntervalTimeCriteria	= TimeInterval*1000.0;
	//_CrtDumpMemoryLeaks();
	return 0;
}

void __stdcall Pulse_SpreadSheet_SaveStatus ( short * IndexFile , 
	  										  long * PulseQty)
{
	if (mPulseExport!=NULL)
	{
	*IndexFile	= mPulseExport->Get_IndexFile();
	*PulseQty	= mPulseExport->Get_PlsCount();
	}
}

void __stdcall Pulse_SpreadSheetStatus ( long * IndexFile , 
										 long * PulseQty)
{
	if (mPulseExport!=NULL)
	{
		*IndexFile	= mPulseExport->Get_IndexFile();
		*PulseQty	= mPulseExport->Get_PlsCount();
		if (*PulseQty<0)
		{
			*PulseQty	= -1;
		}
		if (*IndexFile<0)
		{
			*IndexFile	= -1;
		}
	}
}

long __stdcall Pulse_SpreadSheet_Saved ( bool* lvDone )
{
	if (mPulseExport!=NULL)
	{
		*lvDone	= mPulseExport->Get_SS_Save_Done();
	} else {
		*lvDone	= true; 
	}
	if (*lvDone==true)
	{
		return 1;
	} 
	return 0;
}

long __stdcall Pulse_SpreadSheetDone ( bool* lvDone )
{
	bool	result=false;
	if (mPulseExport!=NULL)
	{
		result = mPulseExport->Get_SSCreation_Done();
	}
	*lvDone = result;
	if (result==true)
	{
		return 1;
	}
	return 0;
}

void __stdcall Pulse_SaveAsStart ( LPSTR strFileName)
{
	//::CreateThread(NULL,0,(LPTHREAD_START_ROUTINE)ThreadSaveSpreadSheet,(LPVOID)(mPulseExport),0,NULL);
}

void __stdcall Pulse_Debug ( short * IndexFile , 
							 long * PulseQty ,
							 LPSTR lvPath )
{
	pulse_import 		pulse_import_cmd;
	pulse_export		pulse_export_cmd;
	__ProjectSt*		Project;

	cFileName	pulse_path(lvPath);
	cFileName	work_space(lvPath);
	cFileName	output(lvPath);
	
	Project					= &mPulseProject.mProject;
	mPulseProject.setMissionPath(&(pulse_path.mName));
	mPulseProject.setWrkSpcPath(&(work_space.mName));
	mPulseProject.setOutputPath(&(output.mName));
	
	Project->FilePwdList_Count	= 0;
	Project->FilesPerWorkSpace	= 4;
	Project->pFilePwdListSt		= NULL;
	
	pulse_import_cmd.import_file(&mPulseProject);
	mPulseProject.mProject.workSheetConfiguration.PulseQtyCriteria = 1500;
	mPulseProject.mProject.workSheetConfiguration.workSheetsPerXlsCount	= 5;
	mPulseProject.create_workspace_byPulse();

	unsigned_long	FileCount;
	unsigned_long	SheetCount;
	unsigned_long	PulseCount;
	unsigned_long	IndexSpread;
	unsigned_long	IndexSheet;
	unsigned_long	IndexPulse;
	double			dPwd[PWD_FIELD_COUNT];

	FileCount		= mPulseProject.GetSpreadFileCount();
	for (IndexSpread=0;IndexSpread<FileCount;IndexSpread++)
	{
		SheetCount	= mPulseProject.GetSheetCount(IndexSpread);
		for (IndexSheet=0;IndexSheet<SheetCount;IndexSheet++)
		{
			PulseCount	= mPulseProject.GetPulseCount(IndexSpread,IndexSheet);
			for (IndexPulse=0;IndexPulse<PulseCount;IndexPulse++)
			{
				mPulseProject.GetPwd(IndexSpread,IndexSheet,IndexPulse,dPwd);
			}
		}
	}
}

void __stdcall Pulse_CreateWorkSpace_BG ( )
{
	::CreateThread(NULL,0,(LPTHREAD_START_ROUTINE)ThreadCreateWorkSpace,(LPVOID)(&mPulseProject),0,NULL);
}

void __stdcall Pulse_CreateWorkSpace ( )
{
	unsigned_long	Index = 0;
	unsigned_long	Count = mPulseProject.LoadErrFileCount();
	//mPulseProject.LoadErrPntLst(Index);
	if (mPulseProject.mProject.workSheetConfiguration.IntervalTimeCriteria > 0)
	{
		mPulseProject.create_workspace_byInterval();
	}
	else
	{
		mPulseProject.create_workspace_byPulse();		
	}
	for (Index=0;Index<Count;Index++)
	{
		mPulseProject.LoadErrPntLst(Index);
		mPulseProject.m_cBookMark.SetPoiterList(&mPulseProject.mProject.mErrPntList);
		mPulseProject.m_cBookMark.UpdateBookMark(mPulseProject.mProject.pProjectFile,
											 &mPulseProject.mProject.mErrPntList,
											 Count);
		mPulseProject.SaveErrPntLst(Index);
		mPulseProject.DestroyErrPntLst();
	}
	mPulseProject.SaveWorkSpace();
}

void __stdcall Pulse_LoadWorkSpace ( )
{
	mPulseProject.LoadWorkSpace();
}

void __stdcall Pulse_DestroyWorkSpace ( )
{
	mPulseProject.DestroyWorkSpace();
}


UINT	ThreadCreateSpreadSheet	( LPVOID pParam )
{
	pulse_export *pXlsExport = (pulse_export *)pParam;
	
	pXlsExport->Run_CreateXls();

	char					lvStr[4096];

	DWORD		lvTickEnd	= GetTickCount();
	sprintf(lvStr,"Saving Time\t%d\t%d\t%d\n",
		pXlsExport->m_TickCount,
		lvTickEnd,
		lvTickEnd-pXlsExport->m_TickCount);
	Pulse_Log_Write(lvStr);
	return 0;
}

UINT	ThreadCreateSpreadSheetOp( LPVOID pParam )
{
	pulse_export *pXlsExport = (pulse_export *)pParam;
	
	pXlsExport->Run_CreateXlsOp();
	
	char					lvStr[4096];
	
	DWORD		lvTickEnd	= GetTickCount();
	sprintf(lvStr,"Saving Time\t%d\t%d\t%d\n",
		pXlsExport->m_TickCount,
		lvTickEnd,
		lvTickEnd-pXlsExport->m_TickCount);
	Pulse_Log_Write(lvStr);
	pXlsExport->Set_SS_Save_Done();
	return 0;
}

UINT	ThreadCreateSpreadProject	( LPVOID pParam )
{
	Pulse_project* pProject = (Pulse_project *)pParam;
	
	mPulseExport->Run_CreateXls(pProject);
	
	return 0;
}


UINT	ThreadSaveSpreadSheet	( LPVOID pParam )
{
	pulse_export *pXlsExport = (pulse_export *)pParam;
	
	pXlsExport->save_xls_file();
	
	return 0;
}

UINT	ThreadPulseImport	( LPVOID pParam )
{
	pulse_import 		pulse_import_cmd;
	
	pulse_import_cmd.m_PoolFilePwdLst	= &mPool_FilePwdLst;
	pulse_import_cmd.m_PoolFilePwd		= &mPool_FilePwd;
	pulse_import_cmd.m_PoolPwdNF		= &mPool_PwdNF;
	pulse_import_cmd.m_PoolPwd			= &mPool_Pwd;
	pulse_import_cmd.m_PoolPwdErr		= &mPool_PwdErr;
	
	//pulse_import_cmd.import_file_optimize((Pulse_project*)pParam);
	pulse_import_cmd.import_file((Pulse_project*)pParam);

	mPulseProject.Set_ImportFileDone();

	return 0;
}

UINT	ThreadCreateWorkSpace	( LPVOID pParam )
{
	Pulse_project *pPulseProject = (Pulse_project *)pParam;
	
	if (pPulseProject->mProject.workSheetConfiguration.IntervalTimeCriteria > 0)
	{
		pPulseProject->create_workspace_byInterval();
	}
	else
	{
		pPulseProject->create_workspace_byPulse();		
	}
	pPulseProject->SaveWorkSpace();
	return 0;
}


void __stdcall Pulse_GetSpreadFileInfo ( long IndexSpread , 
										long * PulseQty ,
										double* TimeIni ,
										double* TimeEnd)
{
	mPulseProject.GetSpreadinfo(IndexSpread,PulseQty,TimeIni,TimeEnd);
}

long __stdcall Pulse_GetSheetCount ( long IndexSpread)
{
	return mPulseProject.GetSheetCount(IndexSpread);
}

long __stdcall Pulse_GetSpreadFileCount ( )
{
	return mPulseProject.GetSpreadFileCount();	
}

void __stdcall Pulse_GetSheetInfo ( long IndexSpread , 
								   long IndexSheet ,
								   long * PulseQty ,
								   double* TimeIni ,
								   double* TimeEnd)
{
	mPulseProject.GetSheetInfo(IndexSpread,IndexSheet,PulseQty,TimeIni,TimeEnd);	
}

void __stdcall Pulse_CreateSheet ( long IndexSpread , 
								  long IndexSheet )
{
	mPulseProject.CreateSheet(IndexSpread,IndexSheet);	
}

void __stdcall Pulse_SaveSpreadSheet ( long IndexSpread )
{
	mPulseProject.SaveSpreadSheet(IndexSpread);	
}

void __stdcall Pulse_GetPwd(long IndexSpread,
							long IndexSheet,
							long IndexPulse, 
							double *Pwd)
{
	mPulseProject.GetPwd(IndexSpread,IndexSheet,IndexPulse,Pwd);
}

void __stdcall Pulse_OutputPath ( LPSTR lvPath )
{
	cFileName	output(lvPath);
	
	mPulseProject.setOutputPath(&(output.mName));
}

void __stdcall Pulse_SetWorkSpacePath ( LPSTR lvPath )
{
	cFileName	work_space(lvPath);
	
	mPulseProject.setWrkSpcPath(&(work_space.mName));
}

void __stdcall Pulse_FilesPerWorkSpace ( long FilesCount )
{
	mPulseProject.mProject.FilesPerWorkSpace	= FilesCount;
}

long __stdcall Pulse_GetSheetPulseCount ( long IndexSpread,long IndexSheet )
{
	return mPulseProject.mProject.pProjectFile->pSpreadFileArray[IndexSpread].pWorkSheetArray[IndexSheet].ul_PulseCount;
}

void __stdcall Pulse_Destroy_All(void)
{
	mPulseProject.DestroyAll();
}

void __stdcall Pulse_GetProjectInfo (long *pPulsesCount, double* pStartTime, double* pStopTime)
{
	if (mPulseProject.mProject.pProjectFile==NULL)
	{
		return;
	}
	*pPulsesCount	= mPulseProject.mProject.pProjectFile->ul_PulseCount;
	*pStartTime		= mPulseProject.mProject.pProjectFile->stPtdIni.d_Time_ms;
	*pStopTime		= mPulseProject.mProject.pProjectFile->stPtdEnd.d_Time_ms;
}

void __stdcall Pulse_GetMissionInfo (LPSTR lvPath,
									 long* FileCount,
									 LPSTR * TimeIni,
									 LPSTR* TimeEnd)
{
	char lvPulsePath[260];
	list<cFileName*>		cListFileName;
	//cFilesList				cFileLst;
	cFileName*				pFileName;
	file_list				c_file_list;
	
#ifndef NDEBUG
	int flag = _CrtSetDbgFlag(_CRTDBG_REPORT_FLAG);
	flag |= _CRTDBG_LEAK_CHECK_DF;
	_CrtSetDbgFlag(flag);
#endif
	
	//_CrtDumpMemoryLeaks();
	strcpy(lvPulsePath,lvPath);
	strcat(lvPulsePath,"\\NORMALES");
	c_file_list.get_file_list(lvPulsePath, &cListFileName);
	*FileCount	= cListFileName.size();

	//Pulse_Log_Write("Pulse_GetMissionInfo Ini");

	if (cListFileName.size())
	{
		//list<cFileName*>::iterator	it;
		//it			= cListFileName.begin();
		pFileName	= cListFileName.front();
		strcpy(*TimeIni,pFileName->mName.FileName+(pFileName->mName.LengthString-23));
		(*TimeIni)[19]	= '\0';

		//it			+= (cListFileName.size()-1);
		pFileName	= cListFileName.back();
		strcpy(*TimeEnd,pFileName->mName.FileName+(pFileName->mName.LengthString-23));
		(*TimeEnd)[19]	= '\0';
	}
	c_file_list.DestroyList(&cListFileName);
	//_CrtDumpMemoryLeaks();
	//Pulse_Log_Write("Pulse_GetMissionInfo End");
	//Pulse_Log_Write("");
}

void __stdcall Pulse_GetMissionName(LPSTR* lsName)
{
	mPulseProject.getMissionName(*lsName);
}

void __stdcall Pulse_SetFieldFormat(long* Order,long* Visible)
{
	mPulseProject.setColumnFormat(Order,Visible);
	mPulseProject.SaveColumnFormat();
}


long __stdcall Pulse_GetIntermediaFileCount(void)
{
	return mPulseProject.get_FilePwdList_Count();
}


void __stdcall Pulse_SetIntermediaFileCount(long lCount)
{
	mPulseProject.set_FilePwdList_Count(lCount);
}


bool __stdcall Pulse_CancelXlsProcess(void)
{
	if (mPulseExport!=NULL)
	{
		return mPulseExport->CancelXlsProcess();
	}
	return true;
}

void __stdcall Pulse_DestroyErrorList(void)
{
	mPulseProject.DestroyErrPntLst();
}

long __stdcall Pulse_GetErrorFileCount(void)
{
	return (long) mPulseProject.m_cBookMark.GetErrorFileCount();
}

long __stdcall Pulse_GetErrorListCount(long Index)
{
	mPulseProject.DestroyErrPntLst();
	mPulseProject.LoadErrPntLst(Index);
	return mPulseProject.mProject.mErrPntList.Count;
}

void __stdcall Pulse_GetErrorFieldCount(long* lCount,long* dCount)
{
	*lCount		= hdrToa + 1;
	*dCount		= hdrEnd - *lCount;
}

void __stdcall Pulse_GetErrorFieldHeader(long Index, LPSTR* lsHeader)
{
	if (Index<hdrEnd)
	{
		strcpy(*lsHeader,GV_ErrorHeader[Index]);
	}
}


void __stdcall Pulse_GetErrorPointer(long Index, long* lPointer,double* dPointer)
{
	__PwdPointerList*	pList	= &mPulseProject.mProject.mErrPntList;
	__PwdPointerSt*		pPtr	= pList->PointerArray;

	if ((pPtr!=NULL) && (Index < pList->Count))
	{
		pPtr			+= Index;
		lPointer[0]		= pPtr->s_Error_Code;
		lPointer[1]		= pPtr->ul_Index_Spread;
		lPointer[2]		= pPtr->ul_Index_Sheet;
		lPointer[3]		= pPtr->ul_Index_Relative;
		lPointer[4]		= pPtr->ul_IndexFile;
		lPointer[5]		= pPtr->ul_Index;
		lPointer[6]		= pPtr->ul_IndexPwd;
		lPointer[7]		= pPtr->ul_Toa;
		dPointer[0]		= pPtr->d_post_d_Toa_us;
		dPointer[1]		= pPtr->d_Rel_Toa_ms;
		dPointer[2]		= pPtr->d_Abs_Toa_s;
	}
}

void __stdcall Pulse_Init_PoolMemory_BG(long FilePwdLst_Size, 
									 long FilePwd_Size, 
									 long PwdNF_Size)
{
	long*		lvArray;

	lvArray = (long*) new long [3];

	lvArray[0] = FilePwdLst_Size;
	lvArray[1] = FilePwd_Size;
	lvArray[2] = PwdNF_Size;

	::CreateThread(NULL,0,(LPTHREAD_START_ROUTINE)ThreadPulseInitPoolMemory,
					(LPVOID)(lvArray),0,NULL);
	
}

void __stdcall Pulse_Init_PoolMemory(long FilePwdLst_Size, 
									 long FilePwd_Size, 
									 long PwdNF_Size)
{
	mPool_FilePwdLst.Pool_CreateTable(FilePwdLst_Size);
	mPool_FilePwd.Pool_CreateTable(FilePwd_Size);
	mPool_PwdNF.Pool_CreateTable(PwdNF_Size);
	mPool_Pwd.Pool_CreateTable(PwdNF_Size);
	mPool_PwdErr.Pool_CreateTable(PwdNF_Size);

	mPulseProject.m_PoolFilePwdLst	= &mPool_FilePwdLst;
	mPulseProject.m_PoolFilePwd		= &mPool_FilePwd;
	mPulseProject.m_PoolPwdNF		= &mPool_PwdNF;
	mPulseProject.m_PoolPwd			= &mPool_Pwd;
	mPulseProject.m_PoolPwdErr		= &mPool_PwdErr;
}

UINT	ThreadPulseInitPoolMemory( LPVOID pParam )
{
	long*		lvArray = (long*)pParam;
	
	Pulse_Init_PoolMemory(lvArray[0],lvArray[1],lvArray[2]);

	return 0;
}

void __stdcall Pulse_GetStructSize(long *FilePwdLst_Size, 
								  long *FilePwd_Size, 
								  long *PwdNF_Size,
								  long *Pwd_Size)
{
	*FilePwdLst_Size	= mPool_FilePwdLst.GetDataSize();
	*FilePwd_Size		= mPool_FilePwd.GetDataSize();
	*PwdNF_Size			= mPool_PwdNF.GetDataSize();
	*Pwd_Size			= mPool_Pwd.GetDataSize();
}

void __stdcall Pulse_Destroy_PoolMemory(void)
{
	mPool_FilePwdLst.Pool_Destroy();
	mPool_FilePwd.Pool_Destroy();
	mPool_PwdNF.Pool_Destroy();
	mPool_Pwd.Pool_Destroy();
	mPool_PwdErr.Pool_Destroy();
}


void __stdcall Pulse_CvtBinXls(long lSheetCount,LPSTR lsFile)
{
	m_DllFnCvtBin(lSheetCount,lsFile);
}

void __stdcall Pulse_Log_Path(LPSTR StrPath)
{
	char		lvFileName[250];
	char		lsTime[250];
	time_t		t_time;
	
	time (&t_time);
	
	sprintf(lsTime,"%s",ctime (&t_time));
	/*012345678901234567890123456789*/
	/*Www Mmm dd hh:mm:ss yyyy*/
	lsTime[7]	= '\0';
	lsTime[19]	= '\0';
	lsTime[24]	= '\0';
	lsTime[10]	= '_';
	lsTime[13]	= '-';
	lsTime[16]	= '-';
	sprintf(lvFileName,"%s\\Dll_Log_%s_%s_%s.txt",StrPath,lsTime+20,lsTime+4,lsTime+8);
	m_Hdl_File	= fopen(lvFileName,"w");
	if (m_Hdl_File!=NULL)
	{
		f_Log_Write_Header();
	}
}

void __stdcall Pulse_Close_Log(void)
{
	if (m_Hdl_File!=NULL)
	{
		fclose(m_Hdl_File);
		m_Hdl_File	= NULL;
	}
}

void Pulse_Log_Write(char* lvData)
{
	if (m_Hdl_File!=NULL)
	{
		char	lsTime[250];
		char	lsData[250];
		time_t	t_time;

		time (&t_time);

		sprintf(lsTime,"%s",ctime (&t_time));
		/*012345678901234567890123456789*/
		/*Www Mmm dd hh:mm:ss yyyy*/
		lsTime[7]	= '\0';
		lsTime[19]	= '\0';
		lsTime[24]	= '\0';
		sprintf(lsData,"%s/%s/%s : %s",lsTime+20,lsTime+4,lsTime+8,lvData);
		f_Log_Write(lsData);
	}
}

void	f_Log_Write(char *lsData)
{
	if (m_Hdl_File!=NULL)
	{
		fprintf(m_Hdl_File,"%s\n",lsData);
	}
}

void	f_Log_Write_Header(void)
{
	char	lsData[250];
	time_t	t_time;
	
	if (m_Hdl_File!=NULL)
	{
		/*App.Title : PulseConvert
		App.Path : C:\Curso Visual\Install\Exe
		Time Start: 20:34:28*/
		f_Log_Write("Dll.Title : Struct.dll");
		f_Log_Write("App.Path : C:\\Curso Visual\\Install\\Lib_Source");
		time (&t_time);
		sprintf(lsData,"Time Start: %s",ctime (&t_time));
		f_Log_Write("");
	}
}

void	Test_Proccess_Map_File(LPSTR StrPath)
{
	ProccessMapFile(StrPath);
}

bool    IsMapFile(char* lsFileName)
{
    char*   p   = lsFileName+(strlen(lsFileName)-4);

    if(strcmp(p,".map")==0)
    {
        return true;
    }
    return false;
}

int     ReadMapFile(char* lsMapFile, list<string> &lsFileList)
{
    FILE*		pFile;
	long		l_FileSize;
	char*		lsFileText;
	char*		pText;
	long		i,j;
	int			iCount		= -1;
	int			iXlsCount	= 0;

    pFile       = fopen(lsMapFile,"rb");
    if(pFile!=NULL)
    {
		fseek( pFile, 0, SEEK_END );
		l_FileSize		=  ftell( pFile );
		rewind(pFile);
		lsFileText		= new char[l_FileSize];
		fread(lsFileText,l_FileSize,sizeof(char),pFile);
		fclose(pFile);
		for (i=0;i<l_FileSize;)
		{
		    pText       = lsFileText + i;
			for (j=i;j<l_FileSize;j++)
			{
				if (lsFileText[j]=='\n' || lsFileText[j]=='\r')
				{
					lsFileText[j]	= '\0';
					if (iCount==-1)
					{
						iCount		= atoi(pText);
						iXlsCount	= 0;
					}
					else
					{
						lsFileList.push_back(pText);
						iXlsCount	++;
					}
					break;
				}
			}
			j++;
			if (lsFileText[j]=='\n' || lsFileText[j]=='\r')
			{
				i	= j+1;
			}
			else
			{
				i = j;
			}
			if (iXlsCount>=iCount)
			{
				break;
			}
		}
    }
    /*else
    {
		lsFileArray = NULL;
    }*/
	return	iCount;
}


void    ProccessMapFile(char* lsMapFile)
{
    list<string>			lsXlsFileArray;
	list<string>::iterator	it;
    int						iFileCount,i;
	string					pStr;

    iFileCount  = ReadMapFile(lsMapFile,lsXlsFileArray);

	::CreateThread(NULL,0,(LPTHREAD_START_ROUTINE)ThreadBinToXls,(LPVOID)(&lsXlsFileArray),0,NULL);

	it			= lsXlsFileArray.begin();
    for(i=0;i<iFileCount;i++)
    {
        //cout << "Archivo " << i << " = " << *it;
		pStr	= *it;
		ConvertBin2Xls(*it);
		it++;
    }
}

void	ConvertBin2Xls(string &strLine)
{
	string		strXls;
	int			SheetCnt;
	string		strSheetCnt;
	string		strMark;
	int			i;
	FILE*		pFile;

	i			= strLine.find(",");
	strSheetCnt	= strLine.substr(0,i);
	strXls		= strLine.substr(i+1);
	SheetCnt	= atoi(strSheetCnt.c_str());
	strMark		= strXls.substr(0,strXls.length() - 3);
	strMark		+= "mrk";
	do 
	{
		pFile	= fopen(strMark.c_str(),"r");
		Sleep(50);
	} while (pFile==NULL);
	fclose(pFile);
	CvtBin2Xls(SheetCnt, strXls.c_str());

}

void	CvtBin2Xls(int SheetCnt,const char* strXls)
{

}

UINT	ThreadBinToXls( LPVOID pParam )
{
	list<string>*			pListFile;
	list<string>::iterator	it;
	string					pStr;
	
	pListFile		= (list<string>*)pParam;

	it			= pListFile->begin();
	for(it=pListFile->begin();it!=pListFile->end();it++)
	{
		//cout << "Archivo " << i << " = " << *it;
		pStr	= *it;
		ConvertBin2Xls(*it);
	}


	return 0;
}



/*it			= lsXlsFileArray.begin();
for(i=0;i<iFileCount;i++)
{
	//cout << "Archivo " << i << " = " << *it;
	pStr	= *it;
	ConvertBin2Xls(*it);
	it++;
}
*/