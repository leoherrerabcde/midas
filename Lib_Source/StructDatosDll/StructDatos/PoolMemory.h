// PoolMemory.h: interface for the CPoolMemory class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_POOLMEMORY_H__58D48BB3_7B0D_4C6B_8ED3_FF62AE88DE1D__INCLUDED_)
#define AFX_POOLMEMORY_H__58D48BB3_7B0D_4C6B_8ED3_FF62AE88DE1D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "pulse_conv_struct_define.h"
#include "PoolTable.h"
#include <stddef.h>
#include <stdio.h>
#include <stdlib.h>

template <class T>
class CPoolMemory : public CPoolTable
{
public:
	CPoolMemory()
	{
		mPool_Memory		= NULL;
	}
	virtual ~CPoolMemory()
	{
		mPool_Memory		= NULL;
		Pool_Destroy();
	}

	void				Pool_CreateTable(long lNewSize)
	{
		CPoolTable::Pool_CreateTable(lNewSize,sizeof(T));
		mPool_Memory	= (T*)CPoolTable::GetPoolTable();
	}

	void				Pool_Destroy(void)
	{
		CPoolTable::Pool_Destroy();		
	}

	T*					Alloc(long Size)
	{
		return 	(T*)CPoolTable::Alloc(Size);
	}
	
	T*					Alloc(void)
	{
		return 	(T*)CPoolTable::Alloc();
	}
	
	void				Free(void* ptn)
	{
		CPoolTable::Free((void*)ptn);
	}

	bool				VerifyAllocated(void* ptn)
	{
		return CPoolTable::VerifyAllocated((void*)ptn);
	}

	long GetSizeTable(void)
	{
		return	CPoolTable::GetSizePoolTable();
	}
	

	long GetDataSize(void)
	{
		return	sizeof(T);
	}

	bool				CheckSpaceAvailable(long Size)
	{
		return CPoolTable::CheckSpaceAvailable(Size);
	}

private:
	T*					mPool_Memory;

};




#endif // !defined(AFX_POOLMEMORY_H__58D48BB3_7B0D_4C6B_8ED3_FF62AE88DE1D__INCLUDED_)
