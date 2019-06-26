/*
 * filelist.cpp
 *
 *  Created on: Sep 17, 2012
 *      Author: lherrera
 */

#include "pulse_conv_struct_define.h"
#include "cFileName.h"
#include "cFilesList.h"
#include <stdio.h>
#include <stdlib.h>
#include <dirent.h>
#include <string.h>
#include <list>
using namespace std;

#include "filelist.h"

void file_list::get_file_list(char * p_ch_Pulse_Path, list<cFileName*>* pFileList)
{
	if (pFileList != NULL)
	{
	    DIR 							* pdir = NULL;
	    struct dirent 					* pent = NULL;
	    cFileName* 						pFileName;
	    unsigned long					ul_TotalSize = 0;
	    unsigned short					us_FileLen;

		pdir			= opendir(p_ch_Pulse_Path);
		if (pdir == NULL){
			return; } 									// end if (pdir == NULL)

		pFileList->clear();
		while ( (pent = readdir (pdir)) ) {
			if (pent == NULL) {
				break;
			} 											// end if (pent == NULL)
			if (strcmp(pent->d_name,".") && strcmp(pent->d_name,"..")) {
				if ((us_FileLen= strlen(pent->d_name))) {
					if (!strcmp(pent->d_name+(us_FileLen-4),".txt")) {
						ul_TotalSize			+= ((unsigned long) pent->d_namlen + 1);
						pFileName				= new cFileName;
						pFileName->SetFileName(pent->d_name,pent->d_namlen);
						pFileList->push_back (pFileName);
					} 	// end if (!strcmp(pent->d_name+(us_FileLen-4),".txt")) {
				}	// end if ((us_FileLen= strlen(pent->d_name))) {
			} 	// end if (strcmp(pent->d_name,".") && strcmp(pent->d_name,".."))
		} //end while
		closedir (pdir);
	}
}

void file_list::get_file_list(list<cFileName*>* pFileList,
							cFilesList* p_cFileList,
							char *pPath,
							long FileQtyPerList,
							long* FileListCount)
{
    list<cFileName*>::iterator 	it,it_i,it_p;
    long					FilesCount;
    long 					l_residuo,i,j,k;
    unsigned long			ul_TotalSize;
    cFileName*				p_cFileName;
    __FileNameSt*			p_FileName;
	__File_List_St* 		p_st_FileList = NULL;
	list<cFileName*>		FileNameList;
	unsigned long			Index;
	char*					p_ch_temp;

    if (pFileList== NULL) { return;}

    FilesCount				= pFileList->size();
	*FileListCount			= FilesCount / FileQtyPerList;
    l_residuo				= FilesCount % FileQtyPerList;
    if (l_residuo) {
    	(*FileListCount)++;
    }

    p_cFileList->SetSize(*FileListCount);
    //it_i = pFileList->begin();

    i = 0;
    j = 0;
    ul_TotalSize			= 0;
    for ( it = pFileList->begin();it!=pFileList->end();it++)
    //for (i=0;i<FileListCount;i++)
    {
    	p_cFileName		= *it;
    	FileNameList.push_back(p_cFileName);
		p_FileName		= &(p_cFileName->mName);
		ul_TotalSize	+= (p_FileName->LengthString + 1);
		j				++;
    	if (j>=FileQtyPerList) {
			p_st_FileList					= p_cFileList->mFileListHead + i;
			p_st_FileList->us_ListCount		= FileQtyPerList;
			p_st_FileList->ul_NameTableSize	= ul_TotalSize;
			p_st_FileList->p_us_NamesLenList= new long [FileQtyPerList];
			p_st_FileList->p_ch_FileList	= new char* [FileQtyPerList];
			p_st_FileList->p_ch_NamesList	= new char [ul_TotalSize];
			p_st_FileList->p_ch_Path		= _strcpy(pPath);
			p_ch_temp						= p_st_FileList->p_ch_NamesList;
			Index							= 0;
			k								= 0;
			for(it_p=FileNameList.begin();it_p!=FileNameList.end();it_p++)
			{
				p_cFileName		= *it_p;
				p_FileName		= &(p_cFileName->mName);
				p_st_FileList->p_ch_FileList[k]		= p_ch_temp;
				p_st_FileList->p_us_NamesLenList[k]	= p_FileName->LengthString ;
				strcpy(p_ch_temp,p_FileName->FileName);
				p_ch_temp		+= (p_FileName->LengthString + 1);
				Index			+= (p_FileName->LengthString + 1);
				k				++;
			}
			FileNameList.clear();
			j					= 0;
			i					++;
			ul_TotalSize		= 0;
    	}
    }
	if (j) {
		p_st_FileList					= p_cFileList->mFileListHead + i;
		p_st_FileList->us_ListCount		= j;
		p_st_FileList->ul_NameTableSize	= ul_TotalSize;	
		p_st_FileList->p_us_NamesLenList= new long [j];
		p_st_FileList->p_ch_FileList	= new char* [j];
		p_st_FileList->p_ch_NamesList	= new char [ul_TotalSize];
		p_st_FileList->p_ch_Path		= new char[strlen(pPath)+1];
		strcpy(p_st_FileList->p_ch_Path,pPath);
		p_ch_temp						= p_st_FileList->p_ch_NamesList;
		Index							= 0;
		k								= 0;
		for(it_p=FileNameList.begin();it_p!=FileNameList.end();it_p++)
		{
			p_cFileName		= *it_p;
			p_FileName		= &(p_cFileName->mName);
			p_st_FileList->p_ch_FileList[k]		= p_ch_temp;
			p_st_FileList->p_us_NamesLenList[k]	= p_FileName->LengthString;
			strcpy(p_ch_temp,p_FileName->FileName);
			p_ch_temp		+= (p_FileName->LengthString + 1);
			Index			+= (p_FileName->LengthString + 1);
			k				++;
		}
		FileNameList.clear();
		j					= 0;
		i					++;
	}
}


char* file_list::_strcpy(char* str)
{
	char*	pStr = new char[strlen(str)+1];
	strcpy(pStr,str);
	return pStr;
}

void file_list::DestroyList(list<cFileName*>* pFileList)
{
    list<cFileName*>::iterator 	it;
	/*cFileName*		pFileName;
	char*			pName;*/

	for (it=pFileList->begin();it!=pFileList->end();it++)
	{
		/*pFileName	= *it;
		pName		= pFileName->mName.FileName ;*/
		delete *it;
	}

}