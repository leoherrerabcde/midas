/*
 * pulseexportexcel.h
 *
 *  Created on: Sep 21, 2012
 *      Author: lherrera
 */

#ifndef PULSEEXPORTEXCEL_H_
#define PULSEEXPORTEXCEL_H_

#include "pulse_conv_struct_define.h"

class pulse_export_excel {
public:
	pulse_export_excel();
	virtual ~pulse_export_excel();

	void 				set_spread_sheet ();
//	bool				export_spread_sheet(char * p_ch_SpreadSheetPath, char *p_ch_SpreadSheetName);
	bool				export_spread_sheet(__File_Pwd_List_St *p_st_File_Pwd_List);
	bool				export_spread_sheet(char * p_ch_SpreadSheetPath, __File_Pwd_List_St *p_st_File_Pwd_List);
	bool				export_spread_sheet(char * p_ch_SpreadSheetName, __File_Pwd_St *p_st_File_Pwd);

private:

	bool				export_spread_sheet(char * p_ch_SpreadSheetName, char * p_ch_SpreadSheetPath, __File_Pwd_St *p_st_File_Pwd );
	char				m_ch_ExcelExtension[5];
};

#endif /* PULSEEXPORTEXCEL_H_ */
