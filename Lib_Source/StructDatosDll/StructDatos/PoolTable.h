// PoolTable.h: interface for the CPoolTable class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_POOLTABLE_H__AA1DE47B_F261_41ED_8CE2_EFDE8FE666BB__INCLUDED_)
#define AFX_POOLTABLE_H__AA1DE47B_F261_41ED_8CE2_EFDE8FE666BB__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "pulse_conv_struct_define.h"
#include "ControlTable.h"
#include <stdio.h>

class CPoolTable : CControlTable 
{
public:
	CPoolTable();
	virtual ~CPoolTable();

	void				Pool_CreateTable(void);
	void				Pool_CreateTable(long lNewSize,
										 long lDataSize);
	void				Pool_CreateTable(long lNewSize);
	void				Pool_SetSize(long lNewSize);
	void				Pool_SetDataSize(long lNewSize);
	void				Pool_ReSize(long nNewSize);
	void				Pool_Destroy(void);

	void*				Alloc(void);
	void*				Alloc(long Size);
	void*				ReAlloc(void* ptn);
	void				Free(void* ptn);
	bool				VerifyAllocated(void * ptn);
	bool				VerifyPointer(void* ptn,long p,long lDataSize);
	bool				CheckSpaceAvailable(long Size);
	void*				GetPoolTable(void);
	void				DumpMemory(char* cPath,char * FileName);
	long				GetSizePoolTable(void);

private:
	typedef struct __pointer
	{
		long			Busy;			// Zero = Free
		long			IndexTable;		// Index to Pool Table
		long			Size;
		long			Index;
		void*			Pointer;
		struct __pointer*		Next;
		struct __pointer*		Prev;
	}_Pointer;
	
	void				Pool_Alloc(void);
	void				Pool_ReAlloc(void);
	void				Set_BusyPointer(_Pointer* pPointer,void* pTable,long nSize);
	void				Set_BusyPointer(_Pointer* pPointer,long nSize);
	void				Set_FreePointer(_Pointer* pPointer,void* pTable,long nSize);

	void				MoveToBusy(_Pointer* p,long nSize);
	void				MoveToFree(_Pointer* p);
	void				RemoveFromFreeWithoutSize(_Pointer* p);
	void				RemoveFromFree(_Pointer* p);
	void				RemoveFromBusy(_Pointer* p);
	void				InsertToFree(_Pointer* pPrev,_Pointer* pNext,_Pointer* p);

	void*				CalcPointerPlusSize(void* p,long nSize);
	long				CalcPointerToPool(void* p);

	void				ReplaceItem(_Pointer* pNew,_Pointer* pOld);
	void				AddToFree(_Pointer* p);
	bool				MergePointers(_Pointer* pPrev,_Pointer* p,_Pointer* pNext);
	void				AppendToFree(_Pointer* p);
	void				AppendToAvailable(_Pointer* p);
	// void				AddSizeToFree(long nSize);
	void				RemoveSizeFromBusy(long nSize);
	void				_Free(void* ptn);
	void*				_Alloc(long Size);
	void				Dump(_Pointer* p,void* ptn,FILE** pFile,char* cName,bool WriteEnable,
							 char* cFunction);
	bool				Debug_Break(long nDataSize,long nDbg_Count,bool bCloseFile);
	void				Set_WriteEnable(void);
	void				Dump_Parameters(FILE** pFile,char* cName,bool WriteEnable,
							 char* cFunction);
	void				Dbg_CloseDump(void);
	bool				DefracPointer(_Pointer* pPointer,CControlTable* pCtrl);
	bool				Dbg_Integridad(char* cFunction);

	_Pointer*			GetAvailable(void){
		if (!mAvailable_Ctrl.Count_Empty)
		{
			if (mDataSize==24)
			{
				mDataSize=24;
			}
			_Pointer*		pNew	= mAvailable_Pnt;

			mAvailable_Pnt			= mAvailable_Pnt->Next;
			mAvailable_Pnt->Prev	= NULL;
			mAvailable_Ctrl.RemoveItem();
			return	pNew;
		} 
		else
		{
			return NULL;
		}
	};

	long				mPool_Size;
	long				mDataSize;
	void*				mPool_Table;
	long*				mTable_Pointers;	//Same Size than mTable
	_Pointer*			mPointer;

	_Pointer*			mFree_Pnt;
	CControlTable		mFree_Ctrl;

	_Pointer*			mBusy_Pnt;
	CControlTable		mBusy_Ctrl;

	_Pointer*			mAvailable_Pnt;
	CControlTable		mAvailable_Ctrl;

	FILE*				mFileFree;
	FILE*				mFileBusy;
	FILE*				mFileParam;

	bool				mFileBusyEnable;
	bool				mFileFreeEnable;
	bool				mFileParamEnable;

	long				Dbg_Count;
};

#endif // !defined(AFX_POOLTABLE_H__AA1DE47B_F261_41ED_8CE2_EFDE8FE666BB__INCLUDED_)
