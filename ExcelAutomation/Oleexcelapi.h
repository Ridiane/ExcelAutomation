#pragma once
#include <Ole2.h>

class Oleexcelapi
{
public:

	Oleexcelapi();		// Constructor
	~Oleexcelapi();		// Deconstructor

	// --< CreateNewInstance : >----------------------------------------------------------------------
	// Create a new Excel instance and get his ID
	// out > pXLApp (IDispatch) = Excel instance's ID
	HRESULT CreateNewInstance(IDispatch **pXLApp);

	// --< GetActiveInstance : >----------------------------------------------------------------------
	// Return an IDispatch interface to a running Excel instance
	IDispatch* GetActiveInstance();

	// --< CloseInstance : >--------------------------------------------------------------------------
	// Close the passed in argument Excel instance.
	// in > pXLApp (IDispatch) = Excel instance
	void CloseInstance(IDispatch *pXLApp);

	// --< SetVisible : >-----------------------------------------------------------------------------
	// Make the Excel instance passed in argument visible or invisible.
	// in > pXLApp (IDispatch) = Excel instance
	// in > arg (int) = 1 = visible / 0 = invisible
	void SetVisible(IDispatch *pXLApp, int arg);

	// --< GetAllWorkbooks : >------------------------------------------------------------------------
	// Return all the workbooks in the Excel instance passed in argument.
	// in > pXLApp (IDispatch) = Excel instance
	IDispatch* GetAllWorkbooks(IDispatch *pXLApp);

	// --< AddWorkbook : >----------------------------------------------------------------------------
	// Add a workbook into the passed workbooks collection, then return it's ID.
	// in > pXLWorbooks (IDispatch) = Targeted Excel instance's workbooks collection
	IDispatch* AddWorkbook(IDispatch *pXLWorkbooks);

	// --< GetActiveWorkbook >------------------------------------------------------------------------
	IDispatch* GetActiveWorkbook(IDispatch *pXLApp);

	// --< GetAllSheet : >----------------------------------------------------------------------------
	IDispatch* GetAllSheets(IDispatch * pXLApp);
	
	// --< GetActiveSheet : >-------------------------------------------------------------------------
	// Return the current active sheet in the targeted Excel instance.
	// in > pXLApp (IDispatch) = Excel instance
	IDispatch* GetActiveSheet(IDispatch *pXLApp);

	// --< GetSheetByName : >-------------------------------------------------------------------------
	// Return the sheet with the name given in the targeted Excel instance.
	// in > pXLApp (IDispatch) = Excel instance
	// in > name (LPOLESTR) = Name of the targeted sheet as it appears in excel
	IDispatch* GetSheetByName(IDispatch * pXLBook, LPOLESTR name);

	// --< AddSheet : >-------------------------------------------------------------------------------
	// Add a new sheet to the given pool of sheets, then return an interface to it.
	// In > pXLSheets (IDispatch) = targeted pool of sheets
	// Out < IDispatch pointer = Interface to the newly added sheet
	IDispatch* AddSheet(IDispatch* pXLSheets);

	// --< SetSheetName : >---------------------------------------------------------------------------
	HRESULT SetSheetName(IDispatch* pXLSheet, LPOLESTR name);

	// --< GetRange : >-------------------------------------------------------------------------------
	// Return the ID of the specified range in the given sheet.
	// in > range (LPOLESTR) = range to be return in excel format (i.e : "A1:B2")
	// in > pxLSheet (IDispatch) = Targeted Excel sheet
	IDispatch* GetRange(LPOLESTR range, IDispatch *pXLSheet);

	// --< SetValueInRange : >------------------------------------------------------------------------
	// Set the given safearrays values in the given cells range.
	// in > val (VARIANT) = safearray containing the desired values. Must be set to VT_ARRAY | VT_VARIANT
	// in > pXLRange (IDispatch) = Targeted cells range
	void SetValueInRange(VARIANT val, IDispatch *pXLRange);

	// --< GetValue : >-------------------------------------------------------------------------------
	// Return the value stored in cells in the given range.
	// in > pXPLRange (IDispatch) = Targeted cells range
	// out < (LPOLESTR) = Content of the cell
	// TODO : 
	//	> Adptative to any types of content in the cell
	VARIANT GetValue(IDispatch *pXLRange);

	// --< SetRangeColor : >--------------------------------------------------------------------------
	// Set the value of the color property in the given range.
	// in > pXLRange (IDispatch) = Targeted cells range
	// in > red, green, blue (int) = RGB code of the desired color
	// out < HRESULT = Success or error code
	HRESULT SetRangeColor(IDispatch *pXLRange, int red, int green, int blue);

private:

	// --< AutoWrap : >------------------------------------------------------------------------------- 
	// Simplifies the code by encapsulating low-level details involved in using IDispatch directly
	HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int argc...);
};

