#include "stdafx.h"
#include "Oleexcelapi.h"
#include <stdio.h>

////// PUBLIC /////////////////////////////////////////////////////////////////////////////////////

Oleexcelapi::Oleexcelapi()
{
}

Oleexcelapi::~Oleexcelapi()
{
}

HRESULT Oleexcelapi::CreateNewInstance(IDispatch **pXLApp)
{
	// Get CLSID
	CLSID clsid;
	HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
	if (FAILED(hr))
		return hr;

	// Start and get IDispatch
	hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)pXLApp);
	return hr;
}

void Oleexcelapi::SetVisible(IDispatch *pXLApp, int arg)
{
	VARIANT result;
	result.vt = VT_I4;
	result.lVal = 1;
	AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXLApp, L"Visible", arg, result);
}

IDispatch* Oleexcelapi::GetAllWorkbooks(IDispatch *pXLApp)
{
	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, pXLApp, L"Workbooks", 0);
	return result.pdispVal;
}

IDispatch* Oleexcelapi::AddWorkbook(IDispatch *pXLWorkbooks)
{
	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, pXLWorkbooks, L"Add", 0);
	return result.pdispVal;
}

IDispatch* Oleexcelapi::GetActiveSheet(IDispatch * pXLApp)
{
	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, pXLApp, L"ActiveSheet", 0);
	return result.pdispVal;
}

void Oleexcelapi::CloseInstance(IDispatch *pXLApp)
{
	AutoWrap(DISPATCH_METHOD, NULL, pXLApp, L"Quit", 0);
}

IDispatch* Oleexcelapi::GetRange(LPOLESTR range, IDispatch *pXLSheet)
{
	VARIANT parm;
	parm.vt = VT_BSTR;
	parm.bstrVal = SysAllocString(range);

	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, pXLSheet, L"Range", 1, parm);
	VariantClear(&parm);

	return result.pdispVal;
}

void Oleexcelapi::SetValueInRange(VARIANT val, IDispatch *pXLRange)
{
	AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXLRange, L"Value", 1, val);
}

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
