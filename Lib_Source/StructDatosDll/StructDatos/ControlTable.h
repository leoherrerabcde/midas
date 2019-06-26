// ControlTable.h: interface for the CControlTable class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CONTROLTABLE_H__9CF3BB4C_1DF2_42C1_B108_EC757DD7AE34__INCLUDED_)
#define AFX_CONTROLTABLE_H__9CF3BB4C_1DF2_42C1_B108_EC757DD7AE34__INCLUDED_


class CControlTable  
{
public:
	CControlTable();
	virtual ~CControlTable();

	bool	AddItem		(void);
	bool	RemoveItem	(void);
	bool	RemoveSize	(long lSize);
	bool	RemoveOnlySize	(long lSize);
	bool	AddSize		(long lNewSize);
	bool	AddOnlySize	(long lNewSize);
	void	SetMax		(long lMaxCount,long lMaxSize);
	void	Set_DataSize(long nDataSize){mDataSize=nDataSize;};
	

	bool	Count_Full;
	bool	Count_Empty;
	bool	Size_Full;
	bool	Size_Empty;

	long	Size;
	long	Count;

	long	Size_Peak;
	long	Count_Peak;

	long	mDataSize;
	
protected:
	
private:
	long	Max_Size;
	long	Max_Count;
};

#endif // !defined(AFX_CONTROLTABLE_H__9CF3BB4C_1DF2_42C1_B108_EC757DD7AE34__INCLUDED_)
