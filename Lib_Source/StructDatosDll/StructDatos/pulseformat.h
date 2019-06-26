/*
 * pulseformat.h
 *
 *  Created on: Sep 19, 2012
 *      Author: lherrera
 */

#ifndef PULSEFORMAT_H_
#define PULSEFORMAT_H_

#include <string.h>
#include "pulse_conv_struct_define.h"

#define LIST_SEPARATE_STR 	";"
#define ABS_TOA_TO_DATE 	86400.0L
//#define OFFSET_TIME			2206731600.0L
//#define OFFSET_TIME			(2206731600.0L + 28*ABS_TOA_TO_DATE)
//#define OFFSET_TIME			(2209150800.0L - 3600.0L)
#define OFFSET_TIME			(2209147200.0L)

class pulse_format {
public:
	pulse_format();
	virtual ~pulse_format();

	double	* format_pwd(__File_Pwd_St *p_st_File_Pwd, unsigned long Index, double * PwdArray);
	char	* format_pwd(char *str_src,__File_Pwd_St *p_st_File_Pwd, unsigned long Index);
	char 	* format_pwd_field(char *str_src,__File_Pwd_St *p_st_File_Pwd, unsigned long Index, short Field);
	double 	format_pwd_field(__File_Pwd_St *p_st_File_Pwd, unsigned long Index, short Field);
	char 	* format_d_Date(char *str_src,double d_Date);
	char	* format_pwd_header(char *str_dst);
	char	* format_pwd_header(char *str_dst, short IndexHeader);
	char	* format_abs_time(char * dest, double d_abs_time);
	void	SetOrder(long* lvOrder){memcpy(mOrder,lvOrder,PWD_FIELD_COUNT*sizeof(long));};
	void	SetVisible(long* lvVisible){memcpy(mVisible,lvVisible,PWD_FIELD_COUNT*sizeof(long));};

private:
	long	mOrder[PWD_FIELD_COUNT];
	long	mVisible[PWD_FIELD_COUNT];
	unsigned_long	m_Offset;
	double	m_d_Offset;
};

#endif /* PULSEFORMAT_H_ */
