// ControlTable.cpp: implementation of the CControlTable class.
//
//////////////////////////////////////////////////////////////////////

#include "ControlTable.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CControlTable::CControlTable()
{
	SetMax(0,0);
}

CControlTable::~CControlTable()
{

}

void CControlTable::SetMax(long lMaxCount, long lMaxSize)
{
	Max_Count	= lMaxCount;
	Max_Size	= lMaxSize;

	Count_Full	= false;
	Size_Full	= false;

	Size_Empty	= false;
	Count_Empty	= false;
	
	Count		= 0;
	Size		= 0;

	Count_Peak	= 0;
	Size_Peak	= 0;
}

bool CControlTable::AddOnlySize	(long lNewSize)
{
	if (Count_Full!=true && Size_Full!=true)
	{
		Size	+= lNewSize;
		
		Size_Empty	= false;
		
		if (Size>=Max_Size)
		{
			Size_Full	= true;
		}
		if (Size>Size_Peak)
		{
			Size_Peak	= Size;
		}
		return true;
	}
	return false;
}

bool CControlTable::AddSize(long lNewSize)
{
	if (Count_Full!=true && Size_Full!=true)
	{
		Count	++;
		Size	+= lNewSize;

		Count_Empty	= false;
		Size_Empty	= false;

		if (Count>=Max_Count)
		{
			Count_Full	= true;
		}
		if (Count>Count_Peak)
		{
			Count_Peak	= Count;
		}
		if (Size>=Max_Size)
		{
			if (mDataSize==24)
			{
				mDataSize=mDataSize;
			}
			Size_Full	= true;
		}
		if (Size>Size_Peak)
		{
			Size_Peak	= Size;
		}
		return true;
	}
	return false;
}

bool CControlTable::RemoveOnlySize(long lSize)
{
	if (Count_Empty!=true && Size_Empty!=true)
	{
		if (Size>=lSize)
		{
			Size	-= lSize;
		}
		
		Size_Full	= false;
		
		if (Size<=0)
		{
			Size_Empty	= true;
		}
		return true;
	}
	return false;
}

bool CControlTable::RemoveSize(long lSize)
{
	if (Count_Empty!=true && Size_Empty!=true)
	{
		Count	--;
		if (Size>=lSize)
		{
			Size	-= lSize;
		}
		
		Count_Full	= false;
		Size_Full	= false;
		
		if (Count<=0)
		{
			if (mDataSize==112)
			{
				mDataSize=mDataSize;
			}
			Count_Empty	= true;
		}
		if (Size<=0)
		{
			Size_Empty	= true;
		}
		return true;
	}
	return false;
}

bool CControlTable::RemoveItem()
{
	if (Count_Empty!=true && Size_Empty!=true)
	{
		Count	--;
		Count_Full	= false;
		
		if (Count<=0)
		{
			if (mDataSize==112)
			{
				mDataSize=mDataSize;
			}
			Count_Empty	= true;
		}
		return true;
	}
	return false;
}

bool CControlTable::AddItem()
{
	if (Count_Full!=true && Size_Full!=true)
	{
		Count	++;
		Count_Empty	= false;
		
		if (Count>=Max_Count)
		{
			Count_Full	= true;
		}
		if (Count>Count_Peak)
		{
			Count_Peak	= Count;
		}
		return true;
	}
	return false;
}
