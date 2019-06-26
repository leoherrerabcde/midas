// TpoReal.h: interface for the CTpoReal class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_TPOREAL_H__FB82AAD0_45DE_4B72_9254_D5E17AA000A9__INCLUDED_)
#define AFX_TPOREAL_H__FB82AAD0_45DE_4B72_9254_D5E17AA000A9__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CTpoReal  
{
public:
	CTpoReal();
	virtual ~CTpoReal();

	unsigned long	GetTickEnlased(void);

private:
	unsigned long	mTickIni;

};

#endif // !defined(AFX_TPOREAL_H__FB82AAD0_45DE_4B72_9254_D5E17AA000A9__INCLUDED_)
