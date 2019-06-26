// BookMark.h: interface for the CBookMark class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_BOOKMARK_H__A691CA5C_C5B9_4344_A429_63DF8260C5A0__INCLUDED_)
#define AFX_BOOKMARK_H__A691CA5C_C5B9_4344_A429_63DF8260C5A0__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "pulse_conv_struct_define.h"

class CBookMark  
{
public:
	CBookMark();
	virtual ~CBookMark();

	void	CleanBookMarks(void);
	void	AddBookMarks	(void);
	void	NextBookMark	(void);
	void	UpdateBookMark	(__PwdIndex* pMarkInfo,__Ptd_St* pPtdSheet);
	void	UpdateBookMark	(__SpreadFileList* pSpreadFileLst,
							 __PwdPointerList* pPntLst,
							 unsigned_long Count);
	void	UpdateBookMark	(__PwdIndex* pMarkInfo,
							 __PwdIndex* pMarkInfoPrevius,
							 unsigned_long IndSpread,
							 unsigned_long IndSheet,
							 unsigned_long IndIndex);
	void	SetPoiterList	(__PwdPointerList* pPntLst);
	unsigned_long GetErrorFileCount(void) {return m_ulErrorFileCount;};
	bool	GetPointOutEnd	(void){return m_PointOutEnd;};
	bool	GetPointOutIni	(void){return m_PointOutIni;};

private:
	__PwdPointerList*	m_PtrList;
	__PwdPointerSt*		m_pArray;
	__PwdPointerSt*		m_BookMark;
	unsigned_long		m_ulErrorFileCount;
	unsigned_long		m_ulCount;
	unsigned_long		m_Index;
	bool				m_PointOutIni;
	bool				m_PointOutEnd;
};

#endif // !defined(AFX_BOOKMARK_H__A691CA5C_C5B9_4344_A429_63DF8260C5A0__INCLUDED_)
