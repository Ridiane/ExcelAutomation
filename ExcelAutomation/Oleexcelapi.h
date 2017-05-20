#pragma once
#include <Ole2.h>
#include <OleAuto.h>

class Oleexcelapi
{
public:
	Oleexcelapi();		// Constructor
	~Oleexcelapi();		// Deconstructor

	HRESULT CreateNewInstance(IDispatch **pXLApp);

	// --< SetVisible : >----------------------------------------------------------------------------
	// Make the Excel instance passed in argument visible or invisible
	// in > pXLApp (IDispatch) = Excel instance
	// in > arg (int) = 1 = visible / 0 = invisible
	void SetVisible(IDispatch *pXLApp, int arg);

	// --< GetAllWorkbooks : >-----------------------------------------------------------------------
	// Return all the workbooks in the Excel instance passed in argument
	// in > pXLApp (IDispatch) = Excel instance
	IDispatch* GetAllWorkbooks(IDispatch *pXLApp);

	// --< AddWorkbook : >-------------------------------------------------------------------------
	IDispatch* AddWorkbook(IDispatch *pXLWorkbooks);

	// --< GetActiveSheet : >----------------------------------------------------------------------
	IDispatch* GetActiveSheet(IDispatch *pXLApp);

	// --< CloseInstance : >-----------------------------------------------------------------------
	void CloseInstance(IDispatch *pXLApp);

	IDispatch* GetRange(LPOLESTR range, IDispatch *pXLSheet);

	// --< SetValueInRange : >---------------------------------------------------------------------
	void SetValueInRange(VARIANT val, IDispatch *pXLRange);

private:

	// --< AutoWrap : >---------------------------------------------------------------------------- 
	// Simplifies the code by encapsulating low-level details involved in using IDispatch directly
	HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int argc...);

protected:

};

