
#include "dirent.h"

#include <windows.h>
#include <tchar.h> 
#include <stdio.h>
//#include <strsafe.h>
#include <string.h>
#pragma comment(lib, "User32.lib")

void DisplayErrorBox(LPTSTR lpszFunction);

DIR				*p_DIR = NULL;
struct dirent	*p_st_dirent;
WIN32_FIND_DATA ffd;
LARGE_INTEGER	filesize;
TCHAR			szDir[MAX_PATH];
size_t			length_of_arg;
HANDLE			hFind = INVALID_HANDLE_VALUE;
DWORD			dwError=0;

DIR*  opendir (const char* pPath)
{
	if (p_DIR!=NULL)
	{
		return NULL;
	} 
	strcpy(szDir,pPath);
	strcat(szDir,"\\*");
	hFind = FindFirstFile(szDir, &ffd);
	
	if (INVALID_HANDLE_VALUE == hFind) 
	{
		return NULL;
	} 
	p_DIR		= new DIR;
	return p_DIR;
}

struct dirent* readdir (DIR*)
{
	if (p_DIR == NULL)
	{
		return NULL;
	} 
	if (p_st_dirent != NULL)
	{
		if(!FindNextFile(hFind, &ffd))
		{
			return NULL;
		}
	}
	else
	{
		p_st_dirent = new struct dirent;
	}
	p_st_dirent->d_ino		= 0;
	p_st_dirent->d_reclen	= 0;
	strcpy	(p_st_dirent->d_name, ffd.cFileName);
	p_st_dirent->d_namlen = (unsigned short)strlen(p_st_dirent->d_name);		
	return p_st_dirent;
}

int  closedir (DIR*)
{
	if (p_DIR	== NULL)
	{
		return 0;
	}
	FindClose(hFind);
	delete p_DIR;
	delete p_st_dirent;
	p_DIR		= NULL;
	p_st_dirent	= NULL;
	return	0;
}

void  rewinddir (DIR*)
{
	
}

long  telldir (DIR*)
{
	return 0;
}

void  seekdir (DIR*, long)
{
	
}



