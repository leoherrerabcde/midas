#include "stdafx.h"



/*
__declspec ( dllexport ) int WINAPI	fnGralFunctions ( void )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	
	return	178;
}
*/

// (int) =	( nByte3 << 24 ) || ( nByte2 << 16 ) || ( nByte1 << 8 ) || ( nByte)
__declspec ( dllexport ) int WINAPI	 Compound_Number(unsigned char nByte0	,
													unsigned char nByte1	,
													unsigned char nByte2	,
													unsigned char nByte3	,
													unsigned char nDigits)
{
	int					lvValue	;
	unsigned char		* pVal;

	pVal				= (unsigned char *) ( & lvValue	);

	* pVal ++			=	nByte0	;

	if ( nDigits > 8 )
	{
		* pVal ++		=	nByte1	;
		if ( nDigits > 16 )
		{
			* pVal ++	=	nByte2	;
			if ( nDigits > 24 )
			{
				* pVal	=	nByte3	;
			}
			else
			{
				* pVal	=	0		;
			}
		}
		else
		{
			* pVal ++	=	0		;
			* pVal		=	0		;
		}
	}
	else
	{
		* pVal ++		=	0		;
		* pVal ++		=	0		;
		* pVal			=	0		;
	}
	return lvValue	;
}

// (int) =	( nByte1 << 8 ) || ( nByte0)
__declspec ( dllexport ) int WINAPI	  Compound_Number16(unsigned char nByte0	,
									unsigned char nByte1	)
{
	int					lvValue	;
	unsigned char		* pVal;
	
	pVal				= (unsigned char *) ( & lvValue	);
	
	* pVal ++			=	nByte0	;
	* pVal ++			=	nByte1	;
	* pVal ++			=	0		;
	* pVal				=	0		;
	return				lvValue	;
}

// (int) =	( nByte2 << 16 ) || ( nByte1 << 8 ) || ( nByte0)
__declspec ( dllexport ) int WINAPI	  Compound_Number24(unsigned char nByte0	,
									unsigned char nByte1	,
									unsigned char nByte2	)
{
	int					lvValue	;
	unsigned char		* pVal;
	
	pVal				= (unsigned char *) ( & lvValue	);
	
	* pVal ++			=	nByte0	;
	* pVal ++			=	nByte1	;
	* pVal ++			=	nByte2	;
	* pVal				=	0		;
	return				lvValue	;
}

// (int) =	( nByte3 << 24 ) || ( nByte2 << 16 ) || ( nByte1 << 8 ) || ( nByte0)
__declspec ( dllexport ) int WINAPI	  Compound_Number32(unsigned char nByte0	,
									unsigned char nByte1	,
									unsigned char nByte2	,
									unsigned char nByte3	)
{
	int					lvValue	;
	unsigned char		* pVal;
	
	pVal				= (unsigned char *) ( & lvValue	);
	
	* pVal ++			=	nByte0	;
	* pVal ++			=	nByte1	;
	* pVal ++			=	nByte2	;
	* pVal				=	nByte3	;
	return				lvValue	;
}

__declspec ( dllexport ) char	Conv_Digit_2_Hex	( unsigned char	lCh)
{
	if ( lCh < 10 )
	{
		return	(char) ( '0' + lCh )	;
	}
	else
	{
		if ( lCh	<	16 )
		{
			return	(char) ( 54 + lCh )	;
		} 
		else
		{
			return		'\0'	;
		}
	}
	return		'\0'	;
}

__declspec ( dllexport ) void WINAPI	Conv_Dec_2_Hex ( int	lDec	, 
														 short	lDigits	,
														 LPSTR	strHex	)
{
	unsigned char		lCh			;
	unsigned char		i			;
	//char				pHex[lDigits+1]		;

	//pHex				= new char[lDigits+1]	;
	
	//pHex[lDigits]		= '\0'	;

	for ( i = 1		; i <= lDigits	; i++ )
	{
		lCh					=	lDec	&	0x0f	;
		strHex[lDigits-i]	=	Conv_Digit_2_Hex ( lCh );
		lDec				>>=	4;
	}
	
	//return				pHex	;
}

__declspec ( dllexport ) void WINAPI	Conv_Byte_2_Str 
									( unsigned char	 * byteArray	, 
									   int lCount						, 
									   LPSTR pStrPrev	)
{
	memcpy			( pStrPrev	,	byteArray	, lCount);
}

__declspec ( dllexport ) int WINAPI	 Conv_Str_2_IdEvent ( LPSTR	pStrEv , unsigned char	lDigits	)
{
	int				lIdEv	;
	//short			lVal	;
	short			* pIdEv	;
	unsigned char	i		;
	unsigned char	lBitRot	;
	unsigned char	lAsc	;
	unsigned char	* pCh	;

	//lIdEv			=	1	;
	pIdEv			= (short *)	& lIdEv	;
	lBitRot			=	0	;
	pCh				= ( unsigned char	* ) pStrEv;

	for ( i = 0 ; i < lDigits	; i ++ )
	{
		lAsc		= ( ( * pCh ++	) - 65 ) & 0x1f	;
		if ( i )
		{
			if ( lBitRot > 7 )
			{
				lBitRot	-=	8;
				* ++ pIdEv	|= ((short)( lAsc )) << ( lBitRot );
			} 
			else
			{
				* pIdEv		|=	((short)( lAsc )) << ( lBitRot );
			}
		}
		else
		{
			 * pIdEv 	=	( lAsc )	;
		}
		lBitRot			+=	5;
	}
	return				lIdEv	;
}


