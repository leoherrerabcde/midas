/*
 * filelist.h
 *
 *  Created on: Sep 17, 2012
 *      Author: lherrera
 */

#include "pulse_conv_struct_define.h"
#include "cFileName.h"
#include "cFilesList.h"

#include <dirent.h>
#include <list>
using namespace std;

#ifndef FILELIST_H_
#define FILELIST_H_

bool file_comparison(struct dirent pentFirst, struct dirent pentSecond);

class file_list {
public:
	file_list();
	virtual ~file_list();

	void 				get_file_list(char * p_ch_Pulse_Path, list<cFileName*>* pFileList);
	void 				get_file_list(list<cFileName*>* pFileList,
								cFilesList* p_cFileList,
								char *pPath,
								long FileQtyPerList,
								long* FileListCount);
	__File_List_St* 	get_file_list(char * p_ch_Pulse_Path);
	void 				get_file_list_op(list<cFileName*>* pFileList,
								cFilesList* p_cFileList,
								char *pPath,
								long FileQtyPerList,
								long* FileListCount);
	void				change_file_extension(char * p_ch_FileName, char *p_ch_New_Extension);
	__File_List_St* 	change_file_extension(__File_List_St * p_File_List_St, char *p_ch_New_Extension);
	void	DestroyList(list<cFileName*>* pFileList);


private:
	char*	_strcpy(char* str);

};

#endif /* FILELIST_H_ */
