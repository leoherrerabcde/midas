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

void Pulse_project::DestroyListIndex	(list<__PwdIndex*>* pListIndex)
{
	list<__PwdIndex*>::iterator	it;
	for (it=pListIndex->begin();it!=pListIndex->end();it++)
	{
		delete *it;
	}
	pListIndex->clear();
}

void Pulse_project::DestroyListWorkSheetBounds	(list<__WorkSheetBounds*>* pWorkSheetBounds)
{
	list<__WorkSheetBounds*>::iterator	it;
	for (it=pWorkSheetBounds->begin();it!=pWorkSheetBounds->end();it++)
	{
		delete *it;
	}
	pWorkSheetBounds->clear();
	
}

void	DestroyListSpreadFile(list<__SpreadFile*>*pSpreadFileList)
{
	list<__SpreadFile*>::iterator	it;
	
	for(it=pSpreadFileList->begin();it!=pSpreadFileList->end();it++)
	{
		delete *it;
	}
	pSpreadFileList->clear();
}

