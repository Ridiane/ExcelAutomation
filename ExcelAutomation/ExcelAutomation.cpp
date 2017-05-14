// ExcelAutomation.cpp : définit le point d'entrée pour l'application console.
//

#include "stdafx.h"
#include <Ole2.h>

HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int argc, char* argv[])
{
	// Variable-argument list
	va_list marker;
	va_start(marker, argc);

	if (!pDisp)
	{
		MessageBox(NULL, "NULL IDispatch passed to AutoWrap()", "Error", 0x10010);
		_exit(0);
	}

	// Variables used...
	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	DISPID dispID;
	HRESULT hr;
	char buf[200];
	char szName[200];

	// Convert down to ANSI
	WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);

	// Get DISPID for name passed...
	hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
	if (FAILED(hr))
	{
		sprintf(buf, "IDispatch::GetIDsOfName(\"%s\") failed w/err 0x%08lx", szName, hr);
		MessageBox(NULL, buf, "AutoWrap()", 0x10010);
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
	hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
	if (FAILED(hr))
	{
		sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%10010", szName, dispID, hr);
		MessageBox(NULL, buf, "AutoWrap()", 0x10010);
		_exit(0);
		return hr;
	}

	va_end(marker);
	delete [] pArgs;
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
		MessageBox(NULL, "CLSIDFromProgID() failed", "Error", 0x10010);
		return -2;
	}

	// Start and get IDispatch
	IDispatch *pXlApp;
	hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pXlApp);
	if (FAILED(hr))
	{
		MessageBox(NULL, "Excel not registered properly", "Error", 0x10010);
		return -2;
	}

	// 
	{
		VARIANT x;
		x.vt = VT_I4;
		x.lVal = 1;
		AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlApp, L"Visible", 1, x); 
		 
	}

}