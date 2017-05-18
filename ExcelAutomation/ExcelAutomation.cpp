// --< ExcelAutomation.cpp : >---------------------------------------------------------------------
//		Define the entry point for the console application.
// ------------------------------------------------------------------------------------------------

#include "stdafx.h"
#include <Ole2.h>

HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int argc...)
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


int main()
{
	// Initialize COM
	CoInitialize(NULL);

	// Get CLSID
	CLSID clsid;
	HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);

	if (FAILED(hr))
	{
		MessageBox(NULL, L"CLSIDFromProgID() failed", L"Error", 0x10010);
		return -2;
	}

	// Start and get IDispatch
	IDispatch *pXlApp;
	hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pXlApp);
	if (FAILED(hr))
	{
		MessageBox(NULL, L"Excel not registered properly", L"Error", 0x10010);
		return -2;
	}

	// Visible
	{
		VARIANT x;
		x.vt = VT_I4;
		x.lVal = 1;
		AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlApp, L"Visible", 1, x);
	}

	// Get all workbooks
	IDispatch *pXlBooks;
	{
		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, L"Workbooks", 0);
		pXlBooks = result.pdispVal;
	}

	// Add a new workbook
	IDispatch *pXlBook;
	{
		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, pXlBooks, L"Add", 0);
		pXlBook = result.pdispVal;
	}

	// Create 15x15 safearray
	VARIANT arr;
	arr.vt = VT_ARRAY | VT_VARIANT;
	{
		SAFEARRAYBOUND sab[2];
		sab[0].lLbound = 1; sab[0].cElements = 15;
		sab[1].lLbound = 1; sab[1].cElements = 15;
		arr.parray = SafeArrayCreate(VT_VARIANT, 2, sab);
	}

	// Fill safearray with some values...
	for (int i = 1; i <= 15; i++)
	{
		for (int j = 1; j <= 15; j++)
		{
			VARIANT tmp;
			tmp.vt = VT_I4;
			tmp.lVal = i*j;
			long indices[] = {i, j};
			SafeArrayPutElement(arr.parray, indices, (void *)&tmp);
		}
	}

	// Get ActiveSheet object
	IDispatch *pXlSheet;
	{
		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, L"ActiveSheet", 0);
		pXlSheet = result.pdispVal;
	}

	// Get Range object for A1:O15
	IDispatch *pXlRange;
	{
		VARIANT parm;
		parm.vt = VT_BSTR;
		parm.bstrVal = SysAllocString(L"A1:O15");

		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, pXlSheet, L"Range", 1, parm);
		VariantClear(&parm);

		pXlRange = result.pdispVal;
	}

	// Set range with our safearray...
	AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlRange, L"Value", 1, arr);

	// Wait for user...
	MessageBox(NULL, L"All Done.", L"Notice", 0x10000);

	// Save
	{
		VARIANT x;
		x.vt = VT_I4;
		x.lVal = 1;
		AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlBook, L"Saved", 1, x);
	}

	// Quit
	AutoWrap(DISPATCH_METHOD, NULL, pXlApp, L"Quit", 0);

	// Release all...
	pXlRange->Release();
	pXlSheet->Release();
	pXlBook->Release();
	pXlBooks->Release();
	pXlApp->Release();
	VariantClear(&arr);

	// Unitialize COM
	CoUninitialize();
}