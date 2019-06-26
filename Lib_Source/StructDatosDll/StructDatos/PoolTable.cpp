// PoolTable.cpp: implementation of the CPoolTable class.
//
//////////////////////////////////////////////////////////////////////

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "PoolTable.h"

/*#include <time.h>
#include <dirent.h>
#include <iostream>*/

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CPoolTable::CPoolTable()
{
	mPool_Size		= 0;
	mDataSize		= 0;
	
	mFree_Pnt		= NULL;
	mBusy_Pnt		= NULL;
	mPool_Table		= NULL;
	mPointer		= NULL;
	mTable_Pointers	= NULL;	//Same Size than mTable

	mFileBusy		= NULL;
	mFileFree		= NULL;

	mFileBusyEnable	= true;
	mFileFreeEnable	= true;

	mAvailable_Pnt	= NULL;

	Dbg_Count		= 0;
}

CPoolTable::~CPoolTable()
{
	if (mPool_Table!=NULL)
	{
		Pool_Destroy();
	}
	Dbg_CloseDump();
}

void CPoolTable::Dbg_CloseDump(void)
{
	if (mFileFree!=NULL)
	{
		fclose(mFileFree);
		mFileFree		=NULL;
		mFileFreeEnable	= false;
	}
	if (mFileBusy!=NULL)
	{
		fclose(mFileBusy);
		mFileBusy		=NULL;
		mFileBusyEnable	= false;
	}
	if (mFileParam!=NULL)
	{
		fclose(mFileParam);
		mFileParam		=NULL;
		mFileParamEnable	= false;
	}
}

void CPoolTable::Pool_CreateTable(void)
{
	long				i;
	_Pointer*			pNext;
	_Pointer*			pPrev;

	if(!mDataSize || !mPool_Size)
	{
		return;
	}
	mPool_Table			= malloc(mPool_Size * mDataSize);
	mPointer			= new _Pointer[mPool_Size];

	mFree_Pnt			= mPointer;
	mFree_Ctrl.SetMax(mPool_Size,mPool_Size);
	mFree_Ctrl.AddSize(mPool_Size);
	mFree_Ctrl.mDataSize=mDataSize;
	mFree_Ctrl.Set_DataSize(mDataSize);

	mBusy_Pnt			= NULL;
	mBusy_Ctrl.SetMax(mPool_Size,mPool_Size);
	mBusy_Ctrl.mDataSize=mDataSize;

	pNext				= mPointer;
	mAvailable_Ctrl.SetMax(mPool_Size-1,mPool_Size-1);
	mAvailable_Ctrl.mDataSize=mDataSize;
	mAvailable_Pnt		= mPointer+1;
	pPrev				= NULL;
	for (i=0;i<mPool_Size;i++)
	{
		pNext->Prev			= pPrev;
		pNext->Index		= i;
		pNext->Pointer		= NULL;
		pNext->Size			= 0;
		pNext->Busy			= 0;
		pNext->IndexTable	= -1;
		if (pPrev!=NULL)
		{
			pPrev->Next		= pNext;
		}
		pPrev				= pNext;
		pNext++;
		if (i)
		{
			mAvailable_Ctrl.AddItem();
		}
	}
	mFree_Pnt->Next			= NULL;
	mAvailable_Pnt->Prev	= NULL;

	pPrev->Next			= NULL;

	Set_FreePointer(mFree_Pnt,mPool_Table,mPool_Size);
}

void CPoolTable::Pool_CreateTable(long lNewSize)
{
	Pool_SetSize(lNewSize);
	Pool_CreateTable();
}

void CPoolTable::Pool_CreateTable(long lNewSize,long lDataSize)
{
	Pool_SetDataSize(lDataSize);
	Pool_SetSize(lNewSize);
	Pool_CreateTable();
}

void CPoolTable::Set_BusyPointer(_Pointer* pPointer,void* pTable,long nSize)
{
	pPointer->Pointer			= pTable;
	pPointer->Size				= nSize;
	pPointer->Busy				= 1;
	pPointer->IndexTable		= CalcPointerToPool(pPointer);
}

void CPoolTable::Set_BusyPointer(_Pointer* pPointer,long nSize)
{
	pPointer->Size				= nSize;
	pPointer->Busy				= 1;
}

void CPoolTable::Set_FreePointer(_Pointer* pPointer,void* pTable,long nSize)
{
	pPointer->Pointer			= pTable;
	pPointer->Size				= nSize;
	pPointer->Busy				= 0;
	pPointer->IndexTable		= CalcPointerToPool(pTable);
}

void CPoolTable::Pool_SetSize(long lNewSize)
{
	mPool_Size		= lNewSize;
}

void CPoolTable::Pool_SetDataSize(long lNewSize)
{
	mDataSize		= lNewSize;
	Set_WriteEnable();
}

void CPoolTable::Pool_ReSize(long nNewSize)
{
	
}
void CPoolTable::Pool_Destroy(void)
{
	if (mTable_Pointers!=NULL)
	{
		delete []mTable_Pointers;
		mTable_Pointers	= NULL;
	}
	if(mPool_Table!=NULL)
	{
		delete []mPool_Table;
		mPool_Table		= NULL;
	}
	if (mPointer!=NULL)
	{
		delete []mPointer;
		mPointer		= NULL;
	}
	mDataSize			= 0;
	mPool_Size			= 0;
}

void* CPoolTable::Alloc(void)
{
	return Alloc(1);	
}

void* CPoolTable::Alloc(long Size)
{	
	void*	ret =	_Alloc(Size);
	/*Dump(mFree_Pnt,ret,&mFileFree,"Free",mFileFreeEnable,"Alloc");
	Dump(mBusy_Pnt,ret,&mFileBusy,"Busy",mFileBusyEnable,"Alloc");
	Dump_Parameters(&mFileParam,"Param_112",mFileParamEnable,"Alloc");*/
	//Debug_Break(24,1,false);0x04e94578
	// VerifyPointer(ret,0x01a47510,112);
	
	// Dbg_Integridad("Alloc");
	return ret;
}
void* CPoolTable::_Alloc(long Size)
{
	_Pointer*	pPnt;

	Dbg_Count++;
	//Debug_Break(112,10,false);
// 	if (mDataSize==24)
// 	{
// 		mDataSize=mDataSize;
// 	}
	
	if (!mFree_Ctrl.Count)
	{
		return NULL;
	}
	pPnt		= mFree_Pnt;
	while(pPnt!=NULL && pPnt->Size<Size)
	{
		pPnt	= pPnt->Next;
	}
	if (pPnt==NULL)
	{
		Dbg_CloseDump();
		return NULL;
	}

	if (pPnt->Size>Size)
	{
		_Pointer*	pNew	= GetAvailable();
		Set_FreePointer(pNew,
						CalcPointerPlusSize(pPnt->Pointer,Size),
						pPnt->Size-Size);
		ReplaceItem(pNew,pPnt);
		mFree_Ctrl.RemoveOnlySize(Size);
		if (pPnt==mFree_Pnt)
		{
			mFree_Pnt	= pNew;
		}
	} 
	else
	{
		RemoveFromFree(pPnt);
	}
	MoveToBusy(pPnt,Size);

	return pPnt->Pointer;
}

void CPoolTable::MoveToBusy(_Pointer* p,long nSize)
{
	if (mBusy_Pnt!=NULL)
	{
		mBusy_Pnt->Next		= p;
	}
	p->Prev				= mBusy_Pnt;
	p->Next				= NULL;
	p->Size				= nSize;
	p->Busy				= 1;

	mBusy_Pnt			= p;
	mBusy_Ctrl.AddSize(nSize);
	// Debug_Break(24,430,true);
}

void CPoolTable::AddToFree(_Pointer* p)
{
	_Pointer*	pNext	= mFree_Pnt;
	_Pointer*	pPrev	= NULL;

	p->Busy			= 0;
	// Debug_Break(112,5,false);
	while (pNext!=NULL)
	{
		if (pNext->Pointer>p->Pointer)
		{
			if (MergePointers(pPrev,p,pNext)!=true)
			{
				InsertToFree(pPrev,pNext,p);
			} 
			break;
		} 
		pPrev	= pNext;
		pNext	= pNext->Next;
	}
}
void CPoolTable::MoveToFree(_Pointer* p)
{
	/*_Pointer*	pNext;
	_Pointer*	pPrev;
	long		nSize = p->Size;

	// Debug_Break(112,14,false);

	RemoveFromBusy(p);
	if (mDataSize==112)
	{
		mDataSize=mDataSize;
	}
	AddToFree(p);

	if (mBusy_Ctrl.Count_Empty)
	{
		pNext	= pNext;
	}
	if (mBusy_Ctrl.Size_Empty)
	{
		pNext	= pNext;
	}*/
}

void* CPoolTable::ReAlloc(void* ptn)
{
	return ptn;	
}

void CPoolTable::Free(void* ptn)
{
	Dbg_Count++;
// 	VerifyPointer(ptn,0x04e94578,24);
// 	VerifyPointer(ptn,0x01a47510,112);

	if (DefracPointer(mFree_Pnt,&mFree_Ctrl)==true)
	{
		mDataSize=mDataSize;
	}
// 	Debug_Break(24,160,false);
	_Free(ptn);
	// Dbg_Integridad("Free");
	if (DefracPointer(mFree_Pnt,&mFree_Ctrl)==true)
	{
		mDataSize=mDataSize;
	}
}

void CPoolTable::_Free(void* ptn)
{
	_Pointer*		p	= mBusy_Pnt;
	
// 	if (mDataSize==24)
// 	{
// 		mDataSize=mDataSize;
// 	}
	
	// Debug_Break(112,66,false);
	if (ptn==NULL)
	{
		return;
	}
	
	while(p!=NULL)
	{
		if (p->Pointer == ptn)
		{
			 // Debug_Break(24,160,false);
			if (mDataSize==112)
			{
				mDataSize=mDataSize;
			}
			//MoveToFree(p);
			RemoveFromBusy(p);
			AddToFree(p);
			return;
		} 
		else
		{
			p			= p->Prev;
		}
	}

	// Verify is pnt already free
	p			= mFree_Pnt;
	while(p!=NULL)
	{
		if (p->Pointer == ptn)
		{
			return;
		} 
		else
		{
			p			= p->Next;
		}
	}
	p=p;
}

long CPoolTable::GetSizePoolTable(void)
{
	return mPool_Size;
}

void* CPoolTable::GetPoolTable(void)
{
	return mPool_Table;
}

void* CPoolTable::CalcPointerPlusSize(void* p,long nSize)
{
	return (void*)((long)p + nSize*mDataSize);
}

long CPoolTable::CalcPointerToPool(void* p)
{
	return ((long)p - (long)mPool_Table)/mDataSize;
}

bool CPoolTable::MergePointers(_Pointer* pPrev,_Pointer* p,_Pointer* pNext)
{
	//bool		result	= false;
	void*		p_Consecutive;

	// Debug_Break(112,13,false);

	if (pPrev!=NULL)
	{
		p_Consecutive		= CalcPointerPlusSize(pPrev->Pointer,pPrev->Size);
		if (p_Consecutive==p->Pointer)
		{
			pPrev->Size		+= p->Size;
			mFree_Ctrl.AddOnlySize(p->Size);

			AppendToAvailable(p);

			p_Consecutive	= CalcPointerPlusSize(pPrev->Pointer,pPrev->Size);
			if (p_Consecutive==pNext->Pointer)
			{
				pPrev->Size			+= pNext->Size;
				//RemoveFromFree(pNext);
				RemoveFromFreeWithoutSize(pNext);
				AppendToAvailable(pNext);
			}
			return true;
		}
	}

	if (pNext!=NULL)
	{
		p_Consecutive	= CalcPointerPlusSize(p->Pointer,p->Size);
		if (p_Consecutive==pNext->Pointer)
		{
			pNext->Size			+= p->Size;
			pNext->Pointer		= p->Pointer;
			pNext->IndexTable	= p->IndexTable;
			mFree_Ctrl.AddOnlySize(p->Size);
			AppendToAvailable(p);
			return true;
		}
	} 
	else
	{
		mFree_Pnt		= p;
		mFree_Pnt->Busy	= 0;
		mFree_Pnt->Next	= NULL;
		mFree_Pnt->Prev	= NULL;
		mFree_Ctrl.AddSize(mFree_Pnt->Size);
	}
	return false;
}

void CPoolTable::AppendToFree(_Pointer* p)
{
	// Debug_Break(112,13,false);
	p->Busy			= 0;
	if (mFree_Pnt==NULL)
	{
		mFree_Pnt		= p;
		mFree_Pnt->Next	= NULL;
		mFree_Pnt->Prev	= NULL;
		mFree_Ctrl.AddSize(mFree_Pnt->Size);
	} 
	else
	{
		_Pointer*	pNext;
		_Pointer*	pPrev	= NULL;

		for(pNext=mFree_Pnt;pNext!=NULL;pNext=pNext->Next)
		{
			if (!pNext->Size)
			{
				pNext->Prev		= p;
				p->Next			= pNext;
				p->Prev			= pPrev;
				if (pNext==mFree_Pnt)
				{
					mFree_Pnt	= p;
				} 
				else
				{
					pPrev->Next	= p;
				}
				mFree_Ctrl.AddSize(p->Size);
				p->Busy			= 0;
				p->Size			= 0;
				return;
			}
			pPrev	= pNext;
		}
		if (pPrev!=NULL)
		{
			pPrev->Next		= p;
			p->Prev			= pPrev;
			p->Next			= NULL;
			mFree_Ctrl.AddSize(p->Size);
			p->Busy			= 0;
			p->Size			= 0;
		}
	}

}

void CPoolTable::RemoveSizeFromBusy(long nSize)
{

}

void CPoolTable::Dump_Parameters(FILE** pFile,char* cName,bool WriteEnable,char* cFunction)
{
	if (WriteEnable==false)
	{
		return;
	}
	char				FileName[200];
	if (cFunction!=NULL)
	{
		/*sprintf(FileName,
			"Z:\\Curso Visual\\Download\\pulse_convert\\Pulsos Curso\\CASE_004\\Dump_Parameters_%d_%s.txt",
			mDataSize,cName);*/
		sprintf(FileName,
			"%s_%d_%s.txt",
			cName,mDataSize,cFunction);
	}
	else
	{
		/*sprintf(FileName,
			"Z:\\Curso Visual\\Download\\pulse_convert\\Pulsos Curso\\CASE_004\\Dump_Parameters_%d.txt",
			mDataSize);*/
		sprintf(FileName,
			"%s_%d.txt",
			cName,mDataSize);
	}
	if (*pFile==NULL)
	{
		*pFile		= fopen(FileName,"w");
		if (pFile==NULL)
		{
			return;
		}
					 // 1 2 3 4 5 6 7 8 9 0 1 2 3 4 5 6 7 8 9 0
		fprintf(*pFile,"\t\t%s\t\t\t\t\t\t%s\t\t\t\t\t\t%s\n",
			"mFree",
			"mBusy",
			"mAvailable");
					  // 1   2   3   4   5   6   7   8   9   0   1   2   3   4   5   6   7   8   9   0   1
		fprintf(*pFile,"%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n",
			"Fn",
			"mSize",
			"Count",			// 3
			"Empty",
			"Full",		// 5
			"Size",
			"Empty",
			"Full",
			"Count",
			"Empty",		// 10
			"Full",
			"Size",
			"Empty",
			"Full",
			"Count",		// 15
			"Empty",
			"Full",
			"Size",
			"Empty",
			"Full",	// 20
			"Dbg_Cnt");
	}
	else
	{
		*pFile		= fopen(FileName,"a");
	}
	if (*pFile!=NULL)
	{
					  // 1   2   3   4   5   6   7   8   9   0   1   2   3   4   5   6   7   8   9   0   1
		fprintf(*pFile,"%s\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\t%d\n",
			cFunction,
			mDataSize,
			mFree_Ctrl.Count,			// 3
			mFree_Ctrl.Count_Empty,
			mFree_Ctrl.Count_Full,		// 5
			mFree_Ctrl.Size,
			mFree_Ctrl.Size_Empty,
			mFree_Ctrl.Size_Full,
			mBusy_Ctrl.Count,
			mBusy_Ctrl.Count_Empty,		// 10
			mBusy_Ctrl.Count_Full,
			mBusy_Ctrl.Size,
			mBusy_Ctrl.Size_Empty,
			mBusy_Ctrl.Size_Full,
			mAvailable_Ctrl.Count,		// 15
			mAvailable_Ctrl.Count_Empty,
			mAvailable_Ctrl.Count_Full,
			mAvailable_Ctrl.Size,
			mAvailable_Ctrl.Size_Empty,
			mAvailable_Ctrl.Size_Full,	// 20
			Dbg_Count);					// 21
		fclose(*pFile);
	}
}

void CPoolTable::Dump(_Pointer* p,void* ptn,FILE** pFile,char* cName,bool WriteEnable,
					  char* cFunction)
{
	if (mDataSize==24)
	{
		mDataSize=24;
	}
	if (WriteEnable==false)
	{
		return;
	}
	if (mDataSize==24)
	{
		mDataSize=24;
	}
	if (p==NULL)
	{
		return;
	}
	bool				bToNext = (p->Next!=NULL);
	char				FileName[200];
	/*sprintf(FileName,"Z:\\Curso Visual\\Download\\pulse_convert\\Pulsos Curso\\CASE_004\\Dump_%d_%s.txt",
		mDataSize,cName);*/
	if (cFunction!=NULL)
	{
		sprintf(FileName,"%s_%d_%s.txt", cName, mDataSize,cFunction);
	} 
	else
	{
		sprintf(FileName,"%s_%d.txt", cName, mDataSize);
	}
	if (*pFile==NULL)
	{
		*pFile		= fopen(FileName,"w");
		if (pFile==NULL)
		{
			return;
		}
		fprintf(*pFile,"%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n",
						"ptn",
						"mSize",
						"Pointer",
						"p->Busy",
						"p->Indx",
						"IndTbl",
						"p->Ptr",
						"p->Size",
						"p->Next",
						"p->Prev",
						"Dbg_Cnt");
	}
	else
	{
		*pFile		= fopen(FileName,"a");
	}
	if (*pFile!=NULL)
	{
		long	TotalMemory	= 0;
		long	lCnt		= 0;

		for (;p!=NULL;)
		{
			if (p->Size)
			{				  //1   2   3   4   5   6   7   8   9  10  11
				fprintf(*pFile,"%x\t%d\t%x\t%x\t%d\t%d\t%x\t%d\t%x\t%x\t%d\n",
					ptn,
					mDataSize,
					p,
					p->Busy,		// 5
					p->Index,
					p->IndexTable,
					p->Pointer,
					p->Size,
					p->Next,
					p->Prev,
					Dbg_Count);
				TotalMemory	+= p->Size;
				lCnt		++;	
				if (bToNext==true)
				{
					p	= p->Next;
				} 
				else
				{
					p	= p->Prev;
				}
			}
			else
			{
				break;
			}
		}
		fprintf(*pFile,"Total Memory: %d\t\tCount: %d\n",TotalMemory,lCnt);
		fclose(*pFile);
	}
}

void CPoolTable::ReplaceItem(_Pointer* pNew,_Pointer* pOld)
{
	_Pointer*		pNext	= pOld->Next;
	_Pointer*		pPrev	= pOld->Prev;

	pNew->Next		= pNext;
	pNew->Prev		= pPrev;
	if(pNext!=NULL)
	{
		pNext->Prev	= pNew;
	}
	if (pPrev!=NULL)
	{
		pPrev->Next	= pNew;
	}
}

bool CPoolTable::Debug_Break(long nDataSize,long nDbg_Count,bool bCloseFile)
{
	if (mDataSize==nDataSize && Dbg_Count == nDbg_Count)
	{
		if (bCloseFile==true)
		{
			DumpMemory("Z:\\Curso Visual\\Download\\pulse_convert\\Pulsos Curso\\CASE_004","Dbg_Pool");
		}
		return true;
	}
	return false;
}

void CPoolTable::Set_WriteEnable(void)
{
	switch (mDataSize)
	{
	case 24:
	case 88:
		mFileFreeEnable	= false;
		mFileParamEnable= false;
		mFileBusyEnable	= false;
		break;
	default:
		mFileFreeEnable	= true;
		mFileParamEnable= true;
		mFileBusyEnable	= true;
	}
}

void CPoolTable::AppendToAvailable(_Pointer* p)
{
	if (mAvailable_Pnt==p)
	{
		p=p;
		return;
	}
	if (mAvailable_Pnt!=NULL)
	{
		mAvailable_Pnt->Prev	= p;
	}
	p->Next					= mAvailable_Pnt;
	p->Prev					= NULL;
	mAvailable_Pnt			= p;
	mAvailable_Ctrl.AddItem();
}

void CPoolTable::RemoveFromFreeWithoutSize(_Pointer* p)
{
	_Pointer*	pNext	= p->Next;
	_Pointer*	pPrev	= p->Prev;
	
	if (pNext!=NULL)
	{
		pNext->Prev		= pPrev;
	}
	if (pPrev!=NULL)
	{
		pPrev->Next		= pNext;
	} 
	else
	{
		mFree_Pnt		= pNext;
	}
	mFree_Ctrl.RemoveItem();
}

void CPoolTable::RemoveFromFree(_Pointer* p)
{
	_Pointer*	pNext	= p->Next;
	_Pointer*	pPrev	= p->Prev;

	if (pNext!=NULL)
	{
		pNext->Prev		= pPrev;
	}
	if (pPrev!=NULL)
	{
		pPrev->Next		= pNext;
	} 
	else
	{
		mFree_Pnt		= pNext;
	}
	mFree_Ctrl.RemoveSize(p->Size);
}

void CPoolTable::RemoveFromBusy(_Pointer* p)
{
	_Pointer*	pNext	= p->Next;
	_Pointer*	pPrev	= p->Prev;
	
	if (pNext!=NULL)
	{
		pNext->Prev		= pPrev;
	}
	else
	{
		mBusy_Pnt		= pPrev;
	}
	if (pPrev!=NULL)
	{
		pPrev->Next		= pNext;
	} 
	p->Busy				= 0;
	if (mDataSize==112)
	{
		mDataSize=mDataSize;
	}
	mBusy_Ctrl.RemoveSize(p->Size);
}

void CPoolTable::DumpMemory(char* cPath,char * FileName)
{
	FILE*		pFile = NULL;
	char		FilePath[250];

	sprintf(FilePath,"%s\\%s",cPath,FileName);

	Dump(mFree_Pnt,NULL,&pFile, FilePath,true,"mFree_Pnt");
	pFile				= NULL;
	Dump(mBusy_Pnt,NULL,&pFile, FilePath,true,"mBusy_Pnt");
	pFile				= NULL;
	Dump_Parameters(&pFile,FilePath,true,"Param");
}

bool CPoolTable::DefracPointer(_Pointer* pPointer,CControlTable* pCtrl)
{
	if (pPointer==NULL)
	{
		return false;
	}
	if (mDataSize==24)
	{
		mDataSize=mDataSize;
	}
	bool			lbNext = (pPointer->Next!=NULL);
	_Pointer*		pPrev	= pPointer;
	_Pointer*		pNext;
	if (lbNext==true)
	{
		pNext		= pPointer->Next;
	} 
	else
	{
		pNext		= pPointer->Prev;
	}
	
	while (pNext!=NULL)
	{
		if (lbNext==true)
		{
			if (pPrev->Pointer>=pNext->Pointer)
			{
				return true;
			} 
			else
			{
				if (pNext->Pointer==CalcPointerPlusSize(pPrev->Pointer,pPrev->Size))
				{
					return true;
				}
				pPrev	= pNext;
				pNext	= pNext->Next;
			}
		}
		else
		{
			if (pNext->Pointer>=pPrev->Pointer)
			{
				return true;
			} 
			else
			{
				pPrev	= pNext;
				pNext	= pNext->Prev;
			}
		}
	}
	return false;
}

bool CPoolTable::Dbg_Integridad(char* cFunction)
{
	long	lSize;

	lSize	= mFree_Ctrl.Size + mBusy_Ctrl.Size;
	if (lSize!=mPool_Size)
	{
		if (cFunction!=NULL)
		{
			char	FileName[100];
			sprintf(FileName,"Dbg_Int_%s",cFunction);
			DumpMemory(
				"Z:\\Curso Visual\\Download\\pulse_convert\\Pulsos Curso\\CASE_004",
				FileName);
		}
		return true;
	} 
	return false;
}

void CPoolTable::InsertToFree(_Pointer* pPrev,_Pointer* pNext,_Pointer* p)
{
	p->Next			= pNext;
	p->Prev			= pPrev;
	if (pPrev!=NULL)
	{
		pPrev->Next	= p;
	}
	else
	{
		mFree_Pnt	= p;
	}
	if (pNext!=NULL)
	{
		pNext->Prev	= p;
	}
	mFree_Ctrl.AddSize(p->Size);
}

bool CPoolTable::VerifyAllocated(void * ptn)
{
	// Debug_Break(112,66,false);
	if (ptn==NULL)
	{
		return false;
	}
	
	_Pointer*		p	= mBusy_Pnt;
	
	while(p!=NULL)
	{
		if (p->Pointer == ptn)
		{
			// Debug_Break(112,5,false);
			return true;
		} 
		else
		{
			p			= p->Prev;
		}
	}
	return false;
}

bool CPoolTable::VerifyPointer(void* ptn,long p,long lDataSize)
{
	if (lDataSize==mDataSize && p!=0 && ptn!=NULL)
	{
		if ((void*)p==ptn)
		{
			return true;
		}
	} 
	return false;
}

bool CPoolTable::CheckSpaceAvailable(long Size)
{
	return true;
}

