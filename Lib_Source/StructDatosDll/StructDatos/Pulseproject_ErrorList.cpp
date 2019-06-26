/*
 * Pulseproject.cpp
 *
 *  Created on: Oct 20, 2012
 *      Author: lherrera
 */

#include "Pulseproject.h"
#include <stdio.h>
#include <string.h>

#include <list>
using namespace std;



void Pulse_project::Set_ErrorPtrLst(list<__PwdPointerSt>* pErrPntLst)
{
	__PwdPointerList*	pPntLst = &mProject.mErrPntList;
	__PwdPointerSt*		pPnt;

	if (pPntLst->PointerArray!=NULL)
	{
		delete []pPntLst->PointerArray;
	}
	if (pErrPntLst->size()>0)
	{
		pPnt					= new __PwdPointerSt[pErrPntLst->size()];
		pPntLst->Count			= pErrPntLst->size();
		pPntLst->PointerArray	= pPnt;
		list<__PwdPointerSt>::iterator	pIt;
		__PwdPointerSt*			pTmp ;
		for (pIt=pErrPntLst->begin();pIt!=pErrPntLst->end();pIt++)
		{
			pTmp = &(*pIt);
			memcpy(pPnt,&(*pIt),sizeof(__PwdPointerSt));
			pPnt++;
		}
	} 
	else
	{
		pPntLst			= NULL;
	}

}

void Pulse_project::Set_ErrorPtrLstNew(__PwdPointerSt* pErrPntLst,unsigned_long ulErrPrtCnt)
{
	__PwdPointerList*	pPntLst = &mProject.mErrPntList;
	//__PwdPointerSt*		pPnt;
	
	pPntLst->Count				= ulErrPrtCnt;
	pPntLst->PointerArray		= pErrPntLst;	
}

void Pulse_project::SetFileNumber(char *FileNumber, unsigned_long Index)
{
	sprintf(FileNumber,"%05d",Index);
}

bool Pulse_project::SaveErrPntLst(void)
{
	return SaveErrPntLst(m_Error_File_Count++);
}

bool Pulse_project::SaveErrPntLst(unsigned_long Index)
{
	FILE*				pFile;
	char				FileName[260];
	__PwdPointerList*	pPwdErrPntLst = &mProject.mErrPntList;
	char				FileNumber[6];

	SetFileNumber(FileNumber,(Index+1));
	sprintf(FileName,"%s\\ErrPntLst%s.bin",mProject.workSpacePath->FileName,FileNumber);
	pFile		= fopen(FileName,"wb");
	if (pFile==NULL)
	{
		return false;
	}

	if(write_in_file(&pPwdErrPntLst->Count,sizeof(unsigned_long),1,pFile)==false)
	{
		return false;
	}
	if (pPwdErrPntLst->Count)
	{
		if(write_in_file(pPwdErrPntLst->PointerArray,
						 sizeof(__PwdPointerSt),
						 pPwdErrPntLst->Count,
						 pFile)==false)
		{
			return false;
		}
	}
	fclose(pFile);
	return true;
}

bool Pulse_project::SaveErrPntLstNew(void)
{
	FILE*				pFile;
	char				FileName[260];
	__PwdPointerList*	pPwdErrPntLst = &mProject.mErrPntList;
	
	sprintf(FileName,"%s\\ErrPntLst.bin",mProject.workSpacePath->FileName);
	pFile		= fopen(FileName,"wb");
	if (pFile==NULL)
	{
		return false;
	}
	
	if(write_in_file(&pPwdErrPntLst->Count,sizeof(unsigned_long),1,pFile)==false)
	{
		return false;
	}
	if (pPwdErrPntLst->Count)
	{
		if(write_in_file(pPwdErrPntLst->PointerArray,
			sizeof(__PwdPointerSt),
			pPwdErrPntLst->Count,
			pFile)==false)
		{
			return false;
		}
	}
	fclose(pFile);
	return true;
}

unsigned_long Pulse_project::LoadErrFileCount(void)
{
	unsigned_long		Count	= 0;
	FILE*				pFile;
	char				FileNumber[6];
	char				FileName[260];
	__PwdPointerList*	pPwdErrPntLst = &mProject.mErrPntList;
	
	do 
	{
		SetFileNumber(FileNumber,(Count+1));
		sprintf(FileName,"%s\\ErrPntLst%s.bin",mProject.workSpacePath->FileName,FileNumber);
		pFile		= fopen(FileName,"rb");
		if (pFile==NULL)
		{
			//fclose(pFile);
			return Count;
		}
		Count++;
		fclose(pFile);
	} while (pFile != NULL);
	return Count;
}

bool Pulse_project::LoadErrPntLst(unsigned_long Numb)
{
	FILE*				pFile;
	char				FileNumber[6];
	char				FileName[260];
	__PwdPointerList*	pPwdErrPntLst = &mProject.mErrPntList;
	
	//sprintf(FileNumber,"%5d",Numb);
	SetFileNumber(FileNumber,(Numb+1));
	sprintf(FileName,"%s\\ErrPntLst%s.bin",mProject.workSpacePath->FileName,FileNumber);
	pFile		= fopen(FileName,"rb");
	if (pFile==NULL)
	{
		return false;
	}

	if(read_in_file(&pPwdErrPntLst->Count,sizeof(unsigned_long),1,pFile)==false)
	{
		return false;
	}
	if (pPwdErrPntLst->Count)
	{
		pPwdErrPntLst->PointerArray	= new __PwdPointerSt[pPwdErrPntLst->Count];
		if(read_in_file(pPwdErrPntLst->PointerArray,
			sizeof(__PwdPointerSt),
			pPwdErrPntLst->Count,
			pFile)==false)
		{
			return false;
		}
	}
	else
	{
		pPwdErrPntLst->PointerArray = NULL;
	}
	fclose(pFile);
	return true;
}

bool Pulse_project::LoadErrPntLstNew(void)
{
	FILE*				pFile;
	char				FileName[260];
	__PwdPointerList*	pPwdErrPntLst = &mProject.mErrPntList;
	
	sprintf(FileName,"%s\\ErrPntLst.bin",mProject.workSpacePath->FileName);
	pFile		= fopen(FileName,"rb");
	if (pFile==NULL)
	{
		return false;
	}
	
	if(read_in_file(&pPwdErrPntLst->Count,sizeof(unsigned_long),1,pFile)==false)
	{
		return false;
	}
	if (pPwdErrPntLst->Count)
	{
		pPwdErrPntLst->PointerArray	= new __PwdPointerSt[pPwdErrPntLst->Count];
		if(read_in_file(pPwdErrPntLst->PointerArray,
			sizeof(__PwdPointerSt),
			pPwdErrPntLst->Count,
			pFile)==false)
		{
			return false;
		}
	}
	else
	{
		pPwdErrPntLst->PointerArray = NULL;
	}
	fclose(pFile);
	return true;
}

