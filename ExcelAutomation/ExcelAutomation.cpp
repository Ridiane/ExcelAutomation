// --< ExcelAutomation.cpp : >---------------------------------------------------------------------
//		Define the entry point for the console application.
// ------------------------------------------------------------------------------------------------

#include "stdafx.h"
#include "Oleexcelapi.h"

int main()
{
	// Initialize COM
	CoInitialize(NULL);
	Oleexcelapi sheet;

	// Get a running Excel instance
	IDispatch *pXLApp =	sheet.GetActiveInstance();

	// Add a new workbook
	IDispatch *pXLWorkbooks = sheet.GetAllWorkbooks(pXLApp);
	IDispatch *pXLBook = sheet.AddWorkbook(pXLWorkbooks);

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
			tmp.vt = VT_BSTR;
			tmp.bstrVal = SysAllocString(L"Ceci est un test");
			long indices[] = { i, j };
			SafeArrayPutElement(arr.parray, indices, (void *)&tmp);
		}
	}

	// Fill the workbook with theses data
	IDispatch *pXLSheet = sheet.GetActiveSheet(pXLApp);
	IDispatch *pXLRange = sheet.GetRange(L"A2:O16", pXLSheet);
	sheet.SetValueInRange(arr, pXLRange);

	// Wait for user to see the result...
	MessageBox(NULL, L"All Done.", L"Notice", 0x10000);

	IDispatch *pXLRange2 = sheet.GetRange(L"A4:A4", pXLSheet);

	MessageBox(NULL, sheet.GetValue(pXLRange2), L"Getting A4", 0x10000);

	// Close 
	sheet.CloseInstance(pXLApp);

	pXLApp->Release();
	pXLWorkbooks->Release();
	pXLBook->Release();
	pXLRange->Release();
	pXLSheet->Release();
	VariantClear(&arr);

	// Unitialize COM
	CoUninitialize();
}