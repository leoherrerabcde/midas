/*
 * pulse_conv_struct_define.h
 *
 *  Created on: Aug 19, 2012
 *      Author: lherrera
 */

#include <stdio.h>

#ifndef PULSE_CONV_STRUCT_DEFINE_H_
#define PULSE_CONV_STRUCT_DEFINE_H_

//#define	PWD_FIELD_COUNT		20	// Rel 1.0.0
//#define	PWD_FIELD_COUNT		22		// Rel 1.0.1
//#define	PWD_FIELD_COUNT		23		// Rel 1.0.2
#define	PWD_FIELD_COUNT		24		// Rel 1.0.2

#define CONVERT_TOA			0

#define MAX_LEN_FILENAME	1024

#define MAX_DTOA_INTERFILE_MAX_MS	8000
#define MAX_DIFF_FILE_TIME_SEC		8	

#define PWD_PW_MAX					65535
#define PWD_PW_MIN					1
#define PWD_AMP_MAX					1000
#define PWD_AMP_MIN					(-9975)
#define PWD_FREC_MAX				182500
#define PWD_FREC_MIN				985

typedef unsigned long				unsigned_long;


#define		COL_UCADJUST			0
#define		COL_UCSTATE				1
#define		COL_SI_AMPLITUD			2
#define		COL_US_AOA				3
#define		COL_US_SYNTH			4
#define		COL_US_PULSEWIDTH		5
#define		COL_UL_IFM				6
#define		COL_UL_TOA				7
#define		COL_UL_TOACORREGIDO		8
#define		COL_D_DATE_S			9
#define		COL_D_PRE__D_TOA_US		10
#define		COL_D_POST_D_TOA_US		11
#define		COL_D_REL_TOA_MS		12
#define		COL_D_REL_S_TOA_MS		13
#define		COL_D_ABS_TOA_S			14
#define		COL_D_FILETIME_S		15
#define		COL_US_FILENUMBER		16
#define		COL_UCROLLOVER			17
#define		COL_UCPULSEDETAIL		18
#define		COL_UCWRAPAROUND		19
#define		COL_UCTOA_ERROR			20
#define		COL_ST_PROCESSERROR		21
#define		COL_UL_REL_INDEX		PWD_FIELD_COUNT-2
#define		COL_UL_INDEX			PWD_FIELD_COUNT-1


//#define		COL_D_POST_D_TOA_MS		12

struct __pwd_st {
	unsigned char 	uc_Adjust;
	unsigned char	uc_State;
	short			si_Amplitud;
	unsigned short	us_Aoa;
	unsigned short	us_Synth;
	unsigned short	us_PulseWidth;
	unsigned short	us_Spare;
	unsigned long	ul_IFM;
	unsigned long	ul_Toa;
	unsigned long	ul_ToaCorregido;
} ;

typedef struct __ErrorFlags{
	unsigned int	Frec_Err:1;
	unsigned int	Amp_Err:1;
	unsigned int	Pw_Err:1;
	unsigned int	Neg_DToa:1;
	unsigned int	Rel_Toa_Err:1;
	unsigned int	Abs_Toa_Err:1;
	unsigned int	FileTimeErr:1;
	unsigned int	Spare:1;
}__ErrorFlag_St;

typedef union {
	unsigned char	ErrorCode;
	__ErrorFlag_St	ErrorFlags;
}__ErrorUnionSt;

struct __pwd_nf_st {
	unsigned long	ul_Index;
	unsigned long	ul_Rel_Index;
	double 			d_Date_s;
	double 			d_pre__d_Toa_us;
	double 			d_post_d_Toa_us;
	double			d_post_d_Toa_ms;
	double 			d_Rel_Toa_ms;
	double 			d_Rel_S_Toa_ms;
	double 			d_Abs_Toa_s;
	double			d_FileTime_s;
	unsigned short	us_FileNumber;
	unsigned char	uc_RollOver;
	unsigned char	uc_PulseDetail;
	unsigned char	uc_WrapAround;
	unsigned char	uc_Toa_Error;
	__ErrorUnionSt  st_ProcessError;
	unsigned char	us_Spare;
} ;

typedef enum{
	__Frec_Error	= 1,
	__Amp_Error		= 2,
	__Pw_Error		= 4,
	__Neg_DToa		= 8,
	__Rel_Toa_Error	= 16,
	__Abs_Toa_Error	= 32,
	__FileTimeDesync= 64
}__VerifyErrorCode;

typedef struct __pwd_pointer{
	unsigned long	ul_Index;
	unsigned long	ul_Index_Relative;
	unsigned long	ul_Toa;
	unsigned long	ul_Index_Spread;
	unsigned long	ul_Index_Sheet;
	unsigned long	ul_Index_Pulse;
	//unsigned long	ul_Index_WorkSpace;
	unsigned long	ul_IndexFilePwdList;
	unsigned long	ul_IndexFilePwd;
	unsigned long	ul_IndexPwd;
	unsigned_long	ul_IndexFile;
	double 			d_post_d_Toa_us;
	double 			d_Rel_Toa_ms;
	double 			d_Abs_Toa_s;
	short			s_Error_Code;
	short			s_Spare1;
}__PwdPointerSt;

/*typedef struct __pwd_pointer_vect{
	__PwdPointerSt*				pPointer;
	struct __pwd_pointer_vect	pNext;
}__PwdPointerVectSt;*/

typedef struct __pwd_pointer_lst{
	unsigned_long		Count;
	__PwdPointerSt*		PointerArray;
}__PwdPointerList;

typedef struct __ErrorDetectionSt {

}__ErrDetectSt;

typedef enum {
	DToa_InterFile_Exceed_Max_Exception=1,
	FileTime_Different_Exceed_Max_Exception=2
} __Toa_Error_Code;

typedef struct __ptd_st {
	unsigned long	ul_Toa;			// TOA
	unsigned_long	ul_Index;
	double			d_postDtoa_us;	// [TOA(i,0) - Toa(i-1,n-1)] / TOA_TO_USEC_FACTOR
	double 			d_D_Toa_ms;		// [TOA(i,0) - Toa(i-1,n-1)] / TOA_TO_MSEC_FACTOR
	double 			d_Time_ms;		// d_Time 	= [TOA(i,j) - Toa(0,0)] / TOA_TO_MSEC_FACTOR
	double 			d_FileTime_s;	// d_Time of FileName
} __Ptd_St;

typedef struct __file_pwd_st {
	long				l_Pulse_Count;
	char 			*	p_ch_FileName;
	struct __pwd_st *	p_st_Pwd;
	struct __pwd_nf_st*	p_st_Pwd_NewFields;
	struct __ptd_st		st_First_Ptd;
	struct __ptd_st		st_Last_Ptd;
	long				l_FileIndex;
	unsigned short		us_FileName_length;
	short				s_WrapAround_Counter;
	unsigned char		uc_OverFlow_Flag;
	unsigned char		uc_Spare;
	short				s_Spare;
} __File_Pwd_St;

typedef struct __file_list_st {
	unsigned short	us_ListCount;
	unsigned short	us_Spare;
	unsigned long	ul_NameTableSize;
	long			*p_us_NamesLenList;
	char 			**p_ch_FileList;
	char			*p_ch_NamesList;
	char			*p_ch_Path;
} __File_List_St;

typedef struct __file_pwd_list_st {
	size_t					sz_Index;
	unsigned short			us_ListCount;
	unsigned short			us_Spare;
	short					s_PathNameLenght;
	short					s_Spare;
	unsigned long			ul_PulseCount;
	double					d_StartTime;
	double					d_EndTime;
	double					d_TotalTime_ms;
	__File_Pwd_St*			p_st_FilePwdArray;
	__File_List_St*			p_st_FileList;
	char*					p_PathName;
} __File_Pwd_List_St;

enum __worksheetConfigType
{
	byPulseFile = 1,
	byPulseQty,
	byTimeInterval
};

typedef struct __pwd_index{
	unsigned short			us_IndexSheet;
	unsigned short			us_IndexFilePwd;
	unsigned long			us_IndexPulse;
	unsigned long			ul_IndexWorkSpace;
	struct __ptd_st			st_Ptd;
	struct __file_pwd_st 	*p_st_FilePwd;
	struct __pwd_st 		*p_st_Pwd;
	struct __pwd_nf_st 		*p_st_Pwd_NewFields;
	unsigned_long			ul_PulseCount;
} __PwdIndex;

typedef struct __worksheet_bounds{
	unsigned long			ul_PulseCount;
	unsigned_long			ul_BoundsCount;
	double					d_Interval;
	__PwdIndex*				p_StartBoundArray;
	__Ptd_St				stPtdIni;
	__Ptd_St				stPtdEnd;
} __WorkSheetBounds;

typedef	struct __spread_file{
	unsigned_long			us_WorkSheetCount;
	unsigned_long			us_PulseCount;
	double					d_Interval;
	__WorkSheetBounds*		pWorkSheetArray;
	__Ptd_St				stPtdIni;
	__Ptd_St				stPtdEnd;
} __SpreadFile;

typedef struct __spread_file_list{
	unsigned_long			us_SpreadFileCount;
	unsigned_long			ul_PulseCount;
	double					d_Interval;
	__SpreadFile*			pSpreadFileArray;
	__Ptd_St				stPtdIni;
	__Ptd_St				stPtdEnd;
} __SpreadFileList;

struct __workSheetConfig{
	long					workSheetsPerXlsCount;
	__worksheetConfigType	workSheetTypeConfiguration;
	double					IntervalTimeCriteria;
	long					PulseQtyCriteria;
	long					ColumnCount;
	long					ColumnOrder[PWD_FIELD_COUNT];
	long					ColumnEnable[PWD_FIELD_COUNT];
};

struct __FileName_st{
	short	LengthString;
	char	*FileName;
};

typedef struct  
{
	unsigned_long			IndexSpread;
	unsigned_long			IndexSheet;
} __Temp_Index;

typedef struct __IndexTable 
{
	unsigned_long			Index;
	unsigned_long			Size;
}_IndexTable;

typedef struct __IndexPointer
{
	unsigned_long			Index;
	void*					Pointer;
}_IndexPointer;

typedef struct __project_config{
	struct __FileName_st	*missionPath;
	struct __FileName_st	*workSpacePath;
	struct __FileName_st	*outputPath;
	struct __FileName_st	*missionName;
	struct __FileName_st	*projectPath;
	__workSheetConfig		workSheetConfiguration;
	long					FilesPerWorkSpace;
	long					FilePwdList_Count;
	long					IndexFilePwdLstSt;
	long					IndexFilePwdSt;
	double					dStartTime;
	double					dStopTime;
	unsigned_long			ul_PulseCount;
	__File_Pwd_List_St*		pFilePwdListSt;
	__File_Pwd_St*			pFilePwdSt;
	unsigned_long			ul_SpreadFilesCount;
	__SpreadFileList*		pProjectFile;
	__Temp_Index			TempIndex;
	__File_List_St*			pFileListSt;
	__PwdPointerList		mErrPntList;
} __ProjectSt;


typedef struct __celda
{
	union{
		double	dValue;
		long	lValue;
	};
}__Celda;

typedef struct  __sheet_file
{
	char			SheetName[250];
	unsigned_long	NameLength;
	unsigned_long	Cols;
	unsigned_long	Rows;
	__Celda*		Celda;
	FILE*			pFile;
}__Sheet_File;

typedef struct __pwd_st 			__Pwd_St;
typedef struct __pwd_nf_st 			__Pwd_NF_St;
//typedef struct __ptd_st 			__Ptd_St;
//typedef struct __file_pwd_st 		__File_Pwd_St;
//typedef struct __file_list_st 		__File_List_St;
//typedef struct __file_pwd_list_st 	__File_Pwd_List_St;
typedef struct __workSheetConfig	__WorkSheetConfig;
typedef struct __FileName_st		__FileNameSt;


enum ErrorHeaderPosition{
	hdrError_Code	= 0,
		hdrXls,
		hdrSheet,
		hdrPulse,
		hdrFile,
		hdrCorrelative,
		hdrIndex,
		hdrToa,
		hdrPost_Dtoa,
		hdrRel_Toa,
		hdrAbs_Toa,
		hdrEnd
};


#endif /* PULSE_CONV_STRUCT_DEFINE_H_ */
