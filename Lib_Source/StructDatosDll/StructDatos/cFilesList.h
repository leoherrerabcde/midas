/*
 * cFilesList.h
 *
 *  Created on: Oct 5, 2012
 *      Author: lherrera
 */

#ifndef CFILESLIST_H_
#define CFILESLIST_H_

#include "pulse_conv_struct_define.h"

class cFilesList {
public:
	cFilesList();
	virtual ~cFilesList();

	void	SetSize(long Count);

	long			mCount;
	__File_List_St	*mFileListHead;
};

#endif /* CFILESLIST_H_ */
