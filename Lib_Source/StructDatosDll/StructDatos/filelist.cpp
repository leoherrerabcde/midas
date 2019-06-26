/*
 * filelist.cpp
 *
 *  Created on: Sep 17, 2012
 *      Author: lherrera
 */

#include "pulse_conv_struct_define.h"
#include <stdio.h>
#include <stdlib.h>
#include <dirent.h>
#include <string.h>
#include <list>
using namespace std;

#include "filelist.h"

file_list::file_list() {
	// TODO Auto-generated constructor stub

}

file_list::~file_list() {
	// TODO Auto-generated destructor stub
}

// Obtener la lista de archivos de pulsos contenidos en 'p_ch_Pulse_Path'
__File_List_St * file_list::get_file_list(char * p_ch_Pulse_Path)
{
	__File_List_St 					* p_st_FileList = NULL;
    DIR 							* pdir = NULL;
    struct dirent 					* pent = NULL;
    struct dirent					* pent_cpy;
    list<struct dirent*> 			stl_list_FileNameList;
    list<struct dirent*>::iterator 	it;
    unsigned long					ul_TotalSize = 0;
    char							* p_ch_NamesTable;
    char 							* p_ch_temp;
    unsigned short					i,us_FileLen;

	pdir			= opendir(p_ch_Pulse_Path);

	if (pdir == NULL)
	{
		return NULL;
	} // end if

	while ( (pent = readdir (pdir)) )
	{
		if (pent == NULL)
		{
			break;
		}
		if (strcmp(pent->d_name,".") && strcmp(pent->d_name,".."))
		{
			if (!(us_FileLen= strlen(pent->d_name)))
			{
				continue;
			}
			if (!strcmp(pent->d_name+(us_FileLen-4),".txt"))
			{
				ul_TotalSize			+= ((unsigned long) pent->d_namlen + 1);
				pent_cpy				= new (struct dirent);
				*pent_cpy				= *pent;
				stl_list_FileNameList.push_back (pent_cpy);
			}
		}
	} //end while

	p_ch_NamesTable		= new char[ul_TotalSize];
	p_st_FileList		= new __File_List_St;

	p_st_FileList->p_ch_NamesList	= p_ch_NamesTable;
	p_st_FileList->ul_NameTableSize	= ul_TotalSize;
	p_st_FileList->us_ListCount		= stl_list_FileNameList.size();
	p_st_FileList->p_ch_FileList	= new char*[p_st_FileList->us_ListCount];
	p_st_FileList->p_us_NamesLenList= new long[p_st_FileList->us_ListCount];

	p_ch_temp			= p_ch_NamesTable;
	i					= 0;

	//stl_list_FileNameList.sort(file_comparison);
	for (it = stl_list_FileNameList.begin(); it != stl_list_FileNameList.end(); it++)
	{
		pent	= *it;
		memcpy(p_ch_temp,pent->d_name,pent->d_namlen+1);
		p_st_FileList->p_ch_FileList[i]		= p_ch_temp;
		p_st_FileList->p_us_NamesLenList[i]	= pent->d_namlen;
		p_ch_temp 		+= pent->d_namlen+1;
		i++;
	}
	closedir (pdir);

	// Creating File List

	return p_st_FileList;
}

void file_list::change_file_extension(char * p_ch_FileName, char *p_ch_New_Extension)
{
	short		s_StartIndex	= strlen(p_ch_FileName) - 3;
	memcpy		(p_ch_FileName + s_StartIndex, p_ch_New_Extension + 1, 3);
}

__File_List_St * file_list::change_file_extension(__File_List_St * p_File_List_St, char *p_ch_New_Extension)
{
	__File_List_St 					* p_st_FileList = NULL;
    unsigned long					ul_TotalSize;
    unsigned short					us_ListCount;
    char 							* p_ch_NamesTableNew;
    char 							* p_ch_temp;
    unsigned short					i;

    ul_TotalSize					= p_File_List_St->ul_NameTableSize;
    us_ListCount					= p_File_List_St->us_ListCount;

	p_st_FileList					= new __File_List_St;

	p_st_FileList->ul_NameTableSize	= ul_TotalSize;
	p_st_FileList->us_ListCount		= us_ListCount;
	p_st_FileList->p_ch_FileList	= new char*[us_ListCount];
	p_st_FileList->p_us_NamesLenList= new long[us_ListCount];
    p_ch_NamesTableNew				= new char[ul_TotalSize];
	p_st_FileList->p_ch_NamesList	= p_ch_NamesTableNew;

	p_ch_temp						= p_ch_NamesTableNew;

	memcpy(p_ch_NamesTableNew,p_File_List_St->p_ch_NamesList,ul_TotalSize);
	memcpy(p_st_FileList->p_us_NamesLenList,p_st_FileList->p_us_NamesLenList,us_ListCount*sizeof(unsigned short));

	for (i = 0; i<us_ListCount ; i++)
	{
		p_st_FileList->p_ch_FileList[i]		= p_ch_temp;
		change_file_extension		(p_ch_temp,p_ch_New_Extension);
		p_ch_temp 		+= p_st_FileList->p_us_NamesLenList[i];
	}

	return p_st_FileList;
}

bool file_comparison(struct dirent pentFirst, struct dirent pentSecond)
{
	if (strcmp(pentFirst.d_name, pentSecond.d_name) > 0)
	{
		return false;
	}
	else
	{
		return true;
	}
}

