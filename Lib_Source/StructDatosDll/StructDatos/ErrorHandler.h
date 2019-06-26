// ErrorHandler.h: interface for the CErrorHandler class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_ERRORHANDLER_H__5D9BAB5F_4263_4C37_B1BC_E0A660D19749__INCLUDED_)
#define AFX_ERRORHANDLER_H__5D9BAB5F_4263_4C37_B1BC_E0A660D19749__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CErrorHandler  
{
public:
	void AddError(int iCodError,char* FileName,char* FnName,char* Description);
	CErrorHandler();
	virtual ~CErrorHandler();

};

#endif // !defined(AFX_ERRORHANDLER_H__5D9BAB5F_4263_4C37_B1BC_E0A660D19749__INCLUDED_)
