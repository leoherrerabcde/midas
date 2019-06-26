// ListToArray.h: interface for the CListToArray class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_LISTTOARRAY_H__860477E9_A0AB_403A_BD39_24EB9EF883F7__INCLUDED_)
#define AFX_LISTTOARRAY_H__860477E9_A0AB_403A_BD39_24EB9EF883F7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <string.h>
#include <list>
using namespace std;

template <class T>
class CListToArray  
{
public:
	CListToArray();
	virtual ~CListToArray();
	CListToArray(list<T*>*pList){
		list<T*>::iterator	it;
		T* pArray	= new T[pList->size()];
		Array		= pArray;
		for(it=pList->begin();it!=pList->end();it++)
		{
			memcpy(pArray++,*it;sizeof(T));
		}
	};
	T* GetArray(void){return Array};

private:
	T* Array=NULL;
};

#endif // !defined(AFX_LISTTOARRAY_H__860477E9_A0AB_403A_BD39_24EB9EF883F7__INCLUDED_)
