// PoolPointer.h: interface for the CPoolPointer class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_POOLPOINTER_H__23A4B5AE_7B97_4823_9D6F_7B2684ECE484__INCLUDED_)
#define AFX_POOLPOINTER_H__23A4B5AE_7B97_4823_9D6F_7B2684ECE484__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "pulse_conv_struct_define.h"

template <class T>
class CPoolPointer  
{
public:
	CPoolPointer();
	virtual ~CPoolPointer();

	typedef struct
	{
		unsigned_long	Index;
		T*				Pointer;
		unsigned_long	Size;
	}_T_element;

	unsigned_long	Size;
	unsigned_long	IndexIni;
	unsigned_long	IndexEnd;
	_T_element*		pIni;
	_T_element*		pEnd;
	unsigned_long	Count;
	bool			Empty;
	bool			Full;
	bool			AlmostFull;
	bool			AlmostEmpty;
	unsigned_long	ThresholdAlmostFull;
	unsigned_long	ThresholdAlmostEmpty;

	void				SetSize(unsigned_long nNewSize);
	void				Destroy(void);

	_T_element*			AddItem(void);
	void				RemoveItem(T* Item);
	void				RemoveItem(unsigned_long Index);

protected:

	typedef struct
	{
		_T_element*		pItem;
		_T_Item*		pNext;
		_T_Item*		pPrev;
	}_T_Item;

	_T_element*		m_Table;
	_T_Item*		m_PointerTable;
	_T_Item*		m_PointerFree;
	_T_Item*		m_PointerBusy;

	unsigned_long CreateTable(void);
	unsigned_long CreateTable(unsigned_long nSize);
};

template <class T> class CPoolPointer::CPoolPointer()
{
	Size				= 0;
	IndexIni			= 0;
	IndexEnd			= 0;
	Count				= 0;

	m_Table				= NULL;
	pIni				= NULL;
	pEnd				= NULL;
	
	Empty				= true;
	Full				= false;
	AlmostFull			= false;
	AlmostEmpty			= true;
	
	ThresholdAlmostFull		= 0;
	ThresholdAlmostEmpty	= 0;
}

template <class T> void	class CPoolPointer::SetSize(unsigned_long nNewSize)
{
	IndexIni			= 0;
	IndexEnd			= 0;
	Count				= 0;
	
	CreateTable(nNewSize);
	pIni				= Table;
	pEnd				= Table;
	
	Empty				= true;
	Full				= false;
	AlmostFull			= false;
	AlmostEmpty			= true;
	
	ThresholdAlmostFull		= 0;
	ThresholdAlmostEmpty	= 0;
}

template <class T> void	class CPoolPointer::Destroy(void)
{
	Size				= 0;
	IndexIni			= 0;
	IndexEnd			= 0;
	Count				= 0;
	
	delete  [] m_Table;
	pIni				= NULL;
	pEnd				= NULL;
	
	Empty				= true;
	Full				= false;
	AlmostFull			= false;
	AlmostEmpty			= true;
	
	ThresholdAlmostFull		= 0;
	ThresholdAlmostEmpty	= 0;
}

template <class T> unsigned_long class CPoolPointer::CreateTable(void)
{
	unsigned_long		i;
	//_T_Item*			pNext = NULL;
	//_T_Item*			pPrev = NULL;

	m_Table				= new _T_element[Size];
	m_List_Pointer		= new _T_Item[Size];

	p_List_Free			= m_List_Pointer;
	p_List_Free->pNext	= NULL;
	p_List_Free->pPrev	= NULL;
	p_List_Free->pItem	= m_Table;
	m_Table->Index		= 0;
	m_Table->Size		= Size;
	m_Table->Pointer	= NULL;
	//pNext				= m_List_Pointer+1;

	/*p_List_Free->pPrev	= NULL;
	pPrev				= m_List_Pointer;
	for(i=1;i<Size;i++)
	{
		pPrev->pNext	= pNext;
		pNext->pPrev	= pPrev;
		pPrev			= pNext;
		pNext++;
	}
	pPrev->pNext		= 0;*/
	m_List_Busy			= NULL;
	p_List_Free->pItem
}

template <class T> unsigned_long class CPoolPointer::CreateTable(unsigned_long nSize)
{
	Size				= nNewSize;
	CreateTable();
}

template <class T> T* class CPoolPointer::Malloc(void)
{
	return Malloc(1);
}

template <class T> T* class CPoolPointer::Malloc(unsigned_long nSize)
{
	
}


#endif // !defined(AFX_POOLPOINTER_H__23A4B5AE_7B97_4823_9D6F_7B2684ECE484__INCLUDED_)
