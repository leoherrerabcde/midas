// BookMark.cpp: implementation of the CBookMark class.
//
//////////////////////////////////////////////////////////////////////

#include "BookMark.h"
#include <stdio.h>

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBookMark::CBookMark()
{
	m_PtrList	= NULL;
	m_pArray	= NULL;
	m_BookMark	= m_BookMark;
	m_ulCount	= 0;
	m_Index		= 0;
	m_PointOutIni	= false;
	m_PointOutEnd	= false;
}

CBookMark::~CBookMark()
{
	
}

void CBookMark::CleanBookMarks(void)
{
	m_Index		= 0;
	if (m_pArray!=NULL)
	{
		m_BookMark	= m_pArray;
	}
}

void CBookMark::AddBookMarks	(void)
{
	
}

void CBookMark::NextBookMark	(void)
{
	if (m_ulCount)
	{
		if (m_PointOutEnd!=true)
		{
			m_Index++;
			if (m_Index>=m_ulCount)
			{
				m_PointOutEnd	= true;
				m_BookMark		= NULL;
			} 
			else
			{
				m_BookMark++;
				m_PointOutIni	= false;
			}
		} 
	}
}

void CBookMark::UpdateBookMark	(__PwdIndex* pMarkInfo,__Ptd_St* pPtdSheet)
{
	if (m_ulCount)
	{
		while (m_BookMark!=NULL)
		{
			if (m_BookMark->ul_Index <= pMarkInfo->st_Ptd.ul_Index)
			{
				m_BookMark->ul_IndexFilePwdList	= pMarkInfo->ul_IndexWorkSpace;
				m_BookMark->ul_IndexFilePwd		= pMarkInfo->us_IndexFilePwd;
				m_BookMark->ul_Index_Relative	= m_BookMark->ul_Index - pPtdSheet->ul_Index;
				//m_BookMark->ul_IndexPwd			= 0;
				//m_BookMark->ul_Index_Sheet		= 0;
				//m_BookMark->ul_Index_WorkSpace	= 0;
				//m_BookMark->ul_Index_PwdIndex	= 0;
				NextBookMark();
			} else {
				break;
			}
		}
	} 
}

void CBookMark::UpdateBookMark	(__PwdIndex* pMarkInfo,
								 __PwdIndex* pMarkInfoPrevius,
								 unsigned_long IndSpread,
								 unsigned_long IndSheet,
								 unsigned_long IndIndex)
{
	if (m_ulCount)
	{
		while (m_BookMark!=NULL)
		{
			if (m_BookMark->ul_IndexFilePwdList!=pMarkInfoPrevius->ul_IndexWorkSpace)
			{
				return;
			} 
			if (m_BookMark->ul_IndexFilePwd!=pMarkInfoPrevius->us_IndexFilePwd)
			{
				return;
			}
			if (m_BookMark->ul_Index<pMarkInfoPrevius->st_Ptd.ul_Index)
			{
				return;
			}
			if (m_BookMark->ul_Index<=pMarkInfo->st_Ptd.ul_Index)
			{
				m_BookMark->ul_Index_Spread		= IndSpread;
				m_BookMark->ul_Index_Sheet		= IndSheet;
				m_BookMark->ul_Index_Pulse		= IndIndex;
				NextBookMark();
			}
			else
			{

				return;
			}
		}
	} 
}


void CBookMark::UpdateBookMark	(__SpreadFileList* pSpreadFileLst,
								 __PwdPointerList* pPntLst,
								 unsigned_long Count)
{
	__SpreadFile*			pSpread;
	__WorkSheetBounds*		pWrkSheet;
	__PwdIndex*				pIndex;
	unsigned_long			IndSpread,IndSheet,IndIndex;

	if (pSpreadFileLst==NULL)
	{
		return;
	}
	pSpread					= pSpreadFileLst->pSpreadFileArray;
	m_ulErrorFileCount		= Count;
	SetPoiterList(pPntLst);
	if (m_ulCount==0)
	{
		return;
	}
	for (IndSpread=0;IndSpread<pSpreadFileLst->us_SpreadFileCount;IndSpread++)
	{
		pWrkSheet			= pSpread->pWorkSheetArray;
		for (IndSheet=0;IndSheet<pSpread->us_WorkSheetCount;IndSheet++)
		{
			pIndex			= pWrkSheet->p_StartBoundArray;
			for (IndIndex=0;IndIndex<pWrkSheet->ul_BoundsCount;IndIndex+=2)
			{
				UpdateBookMark(pIndex+1,pIndex,IndSpread,IndSheet,IndIndex);
				if (m_PointOutEnd==true)
				{
					return;
				}
				pIndex += 2;
			}
			if (m_PointOutEnd==true)
			{
				return;
			}
			pWrkSheet ++;
		}
		if (m_PointOutEnd==true)
		{
			return;
		}
		pSpread ++;
	}

}

void CBookMark::SetPoiterList	(__PwdPointerList* pPntLst)
{
	m_PtrList		= pPntLst;
	m_pArray		= m_PtrList->PointerArray;
	m_ulCount		= m_PtrList->Count;

	m_BookMark		= m_pArray;
	m_Index			= 0;
	m_PointOutIni	= true;
	m_PointOutEnd	= false;
}

