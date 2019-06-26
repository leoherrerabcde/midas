/*
 * cFileName.h
 *
 *  Created on: Oct 5, 2012
 *      Author: lherrera
 */

#ifndef CFILENAME_H_
#define CFILENAME_H_

#include "pulse_conv_struct_define.h"

class cFileName {
public:
	cFileName();
	virtual ~cFileName();

	cFileName(char *FileName);
	cFileName(char *FileName,short FileNameLength);
	void SetFileName(char *FileName);
	void SetFileName(char *FileName,short FileNameLength);

	__FileNameSt mName;
};

#endif /* CFILENAME_H_ */
