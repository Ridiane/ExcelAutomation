#include "stdafx.h"
#include "Oleexcelapi.h"

////// PUBLIC /////////////////////////////////////////////////////////////////////////////////////

Oleexcelapi::Oleexcelapi()
{
}

Oleexcelapi::~Oleexcelapi()
{
}

// --< CreateNewInstance : >-----------------------------------------------------------------------
// Create a new Excel instance and get his ID
// out > pXLApp (IDispatch) = Excel instance's ID
HRESULT Oleexcelapi::CreateNewInstance(IDispatch **pXLApp)
{
	// Get CLSID from registry
	CLSID clsid;
	HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
	if (FAILED(hr))
		return hr;

	// Start and get IDispatch
	hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)pXLApp);
	return hr;
}
// -----------------------------------------------------------------------------------------< ! >--

// --< GetActiveInstance : >-----------------------------------------------------------------------
// Return an IDispatch interface to a running Excel instance
IDispatch* Oleexcelapi::GetActiveInstance()
{
	// Get CLSID from registry
	CLSID clsid;
	HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
	
	if (FAILED(hr))
	{
		// TODO
		return NULL;
	}

	// Get an interface to the running instance, if there is one...
	IUnknown * pUnk;
	hr = GetActiveObject(clsid, NULL, (IUnknown**)&pUnk);

	if (FAILED(hr))
	{
		// TODO
		return NULL;
	}

	// Get the IDispatch for Automation
	IDispatch * pDisp;
	hr = pUnk->QueryInterface(IID_IDispatch, (void**)&pDisp);

	if (FAILED(hr))
	{
		// TODO
		return NULL;
	}

	pUnk->Release();
	return pDisp;
}
// -----------------------------------------------------------------------------------------< ! >--

// --< CloseInstance : >---------------------------------------------------------------------------
// Close the passed in argument Excel instance.
// in > pXLApp (IDispatch) = Excel instance
void Oleexcelapi::CloseInstance(IDispatch *pXLApp)
{
	AutoWrap(DISPATCH_METHOD, NULL, pXLApp, L"Quit", 0);
}
// -----------------------------------------------------------------------------------------< ! >--

// --< SetVisible : >------------------------------------------------------------------------------
// Make the Excel instance passed in argument visible or invisible.
// in > pXLApp (IDispatch) = Excel instance
// in > arg (int) = 1 = visible / 0 = invisible
void Oleexcelapi::SetVisible(IDispatch *pXLApp, int arg)
{
	VARIANT result;
	result.vt = VT_I4;
	result.lVal = arg;
	AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXLApp, L"Visible", 1, result);
}
// -----------------------------------------------------------------------------------------< ! >--

// --< GetAllWorkbooks : >-------------------------------------------------------------------------
// Return all the workbooks in the Excel instance passed in argument.
// in > pXLApp (IDispatch) = Excel instance
IDispatch* Oleexcelapi::GetAllWorkbooks(IDispatch *pXLApp)
{
	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, pXLApp, L"Workbooks", 0);
	return result.pdispVal;
}
// -----------------------------------------------------------------------------------------< ! >--

// --< AddWorkbook : >-----------------------------------------------------------------------------
// Add a workbook into the passed workbooks collection, then return it's ID.
// in > pXLWorbooks (IDispatch) = Targeted Excel instance's workbooks collection
IDispatch* Oleexcelapi::AddWorkbook(IDispatch *pXLWorkbooks)
{
	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, pXLWorkbooks, L"Add", 0);
	return result.pdispVal;
}
// -----------------------------------------------------------------------------------------< ! >--

// --< GetActiveWorkbook >-------------------------------------------------------------------------
IDispatch* Oleexcelapi::GetActiveWorkbook(IDispatch *pXLApp)
{
	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, pXLApp, L"ActiveWorkbook", 0);
	return result.pdispVal;
}
// -----------------------------------------------------------------------------------------< ! >--

// --< GetAllSheet : >-----------------------------------------------------------------------------
IDispatch* Oleexcelapi::GetAllSheets(IDispatch * pXLApp)
{
	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, pXLApp, L"Sheets", 0);
	return result.pdispVal;
}
// -----------------------------------------------------------------------------------------< ! >--

// --< GetActiveSheet : >--------------------------------------------------------------------------
// Return the current active sheet in the targeted Excel instance.
// in > pXLApp (IDispatch) = Excel instance
IDispatch* Oleexcelapi::GetActiveSheet(IDispatch * pXLApp)
{
	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, pXLApp, L"ActiveSheet", 0);
	return result.pdispVal;
}
// -----------------------------------------------------------------------------------------< ! >--

// --< GetSheetByName : >--------------------------------------------------------------------------
// Return the sheet with the name given in the targeted Excel instance.
// in > pXLApp (IDispatch) = Excel instance
// in > name (LPOLESTR) = Name of the targeted sheet as it appears in excel
IDispatch* Oleexcelapi::GetSheetByName(IDispatch * pXLBook, LPOLESTR name)
{
	// Convert range from LPOLESTR to VARIANT to be used as parameters
	VARIANT parm;
	parm.vt = VT_BSTR;
	parm.bstrVal = SysAllocString(name);

	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, pXLBook, L"Sheets", 1, parm);
	return result.pdispVal;
}
// -----------------------------------------------------------------------------------------< ! >--

// --< AddSheet : >--------------------------------------------------------------------------------
// Add a new sheet to the given pool of sheets, then return an interface to it.
// In > pXLSheets (IDispatch) = targeted pool of sheets
// Out < IDispatch pointer = Interface to the newly added sheet
IDispatch* Oleexcelapi::AddSheet(IDispatch* pXLSheets)
{
	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, pXLSheets, L"Add", 0);
	return result.pdispVal;
}
// -----------------------------------------------------------------------------------------< ! >--

// --< SetSheetName : >----------------------------------------------------------------------------
HRESULT Oleexcelapi::SetSheetName(IDispatch* pXLSheet, LPOLESTR name)
{
	// Convert range from LPOLESTR to VARIANT to be used as parameters
	VARIANT parm;
	parm.vt = VT_BSTR;
	parm.bstrVal = SysAllocString(name);

	HRESULT hr = AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXLSheet, L"Name", 1, parm);
	return hr;
}
// -----------------------------------------------------------------------------------------< ! >--

// --< GetRange : >--------------------------------------------------------------------------------
// Return the ID of the specified range in the given sheet.
// in > range (LPOLESTR) = range to be return in excel format (i.e : "A1:B2")
// in > pxLSheet (IDispatch) = Targeted Excel sheet
IDispatch* Oleexcelapi::GetRange(LPOLESTR range, IDispatch *pXLSheet)
{
	// Convert range from LPOLESTR to VARIANT to be used as parameters
	VARIANT parm;
	parm.vt = VT_BSTR;
	parm.bstrVal = SysAllocString(range);

	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, pXLSheet, L"Range", 1, parm);
	VariantClear(&parm);

	return result.pdispVal;
}
// -----------------------------------------------------------------------------------------< ! >--

// --< SetValueInRange : >-------------------------------------------------------------------------
// Set the given safearrays values in the given cells range.
// in > val (VARIANT) = safearray containing the desired values. Must be set to VT_ARRAY | VT_VARIANT
// in > pXLRange (IDispatch) = Targeted cells range
void Oleexcelapi::SetValueInRange(VARIANT val, IDispatch *pXLRange)
{
	AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXLRange, L"Value", 1, val);
}
// -----------------------------------------------------------------------------------------< ! >--

// --< GetValue : >--------------------------------------------------------------------------------
// Return the value stored in cells in the given range.
// in > pXPLRange (IDispatch) = Targeted cells range
// out < (LPOLESTR) = Content of the cell
// TODO : 
//	> Adptative to any types of content in the cell
//  > Array of cells values instead of only one cell
VARIANT Oleexcelapi::GetValue(IDispatch *pXLRange)
{
	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, pXLRange, L"Value", 0);
	return result;
}
// -----------------------------------------------------------------------------------------< ! >--

// --< SetRangeColor : >---------------------------------------------------------------------------
// Set the value of the color property in the given range.
// in > pXLRange (IDispatch) = Targeted cells range
// in > red, green, blue (int) = RGB code of the desired color
// out < HRESULT = Success or error code
HRESULT Oleexcelapi::SetRangeColor(IDispatch *pXLRange, int red, int green, int blue)
{
	// Get the Interior object (https://msdn.microsoft.com/en-us/library/office/ff196598.aspx) from the targeted range
	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, pXLRange, L"Interior", 0);
	IDispatch *pXLInterior = result.pdispVal;

	// Convert colors arguments to a color value then store it as a Variant(int)
	VARIANT vRGB;
	vRGB.vt = VT_I4;
	vRGB.lVal = RGB(red, green, blue);

	HRESULT hr = AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXLInterior, L"Color", 1, vRGB);
	return hr;
}
// -----------------------------------------------------------------------------------------< ! >--

////// PRIVATE ////////////////////////////////////////////////////////////////////////////////////

// --< AutoWrap : >-------------------------------------------------------------------------------- 
// Simplifies the code by encapsulating low-level details involved in using IDispatch directly
HRESULT Oleexcelapi::AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int argc ...)
{
	// Variable-argument list
	va_list marker;
	va_start(marker, argc);

	if (!pDisp)
	{
		MessageBox(NULL, L"NULL IDispatch passed to AutoWrap()", L"Error", 0x10010);
		_exit(0);
	}

	// Variables used...
	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	DISPID dispID;
	HRESULT hr;
	EXCEPINFO *pexcepinfo = new EXCEPINFO;
	char buf[200];
	char szName[200];

	// Convert down to ANSI
	WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);

	// Get DISPID for name passed...
	hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
	if (FAILED(hr))
	{
		sprintf(buf, "IDispatch::GetIDsOfName(\"%s\") failed w/err 0x%08lx", szName, hr);
		wchar_t wbuf[512];
		mbstowcs(wbuf, buf, 512);
		MessageBox(NULL, wbuf, L"AutoWrap()", 0x10010);
		_exit(0);
		return hr;
	}

	// Memory allocation for arguments
	VARIANT *pArgs = new VARIANT[argc + 1];

	// Extract arguments
	for (int i = 0; i < argc; i++)
		pArgs[i] = va_arg(marker, VARIANT);

	// Build DISPPARAMS
	dp.cArgs = argc;
	dp.rgvarg = pArgs;

	// Handle special-case for property-puts
	if (autoType & DISPATCH_PROPERTYPUT)
	{
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispidNamed;
	}

	// Make the call
	hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, pexcepinfo, NULL);
	if (FAILED(hr))
	{
		if (hr == DISP_E_EXCEPTION)
			sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx", szName, dispID, pexcepinfo->scode);
		else
			sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx", szName, dispID, hr);

		wchar_t wbuf[512];
		mbstowcs(wbuf, buf, 512);
		MessageBox(NULL, wbuf, L"AutoWrap()", 0x10010);
		_exit(0);
		return hr;
	}

	va_end(marker);
	delete[] pArgs;
	return hr;
}
// -----------------------------------------------------------------------------------------< ! >--
