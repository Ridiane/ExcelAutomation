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
	IDispatch *pXLApp = sheet.GetActiveInstance();
	if (pXLApp == NULL)
	{
		MessageBox(NULL, L"No Excel instances found.", L"An error occured!", 0x10010);
		return -1;
	}
	sheet.SetVisible(pXLApp, 1);

	// Add a new workbook
	// IDispatch *pXLWorkbooks = sheet.GetAllWorkbooks(pXLApp);
	// IDispatch *pXLBook = sheet.AddWorkbook(pXLWorkbooks);

	// Get the currently active workbook
	IDispatch *pXLBook = sheet.GetActiveWorkbook(pXLApp);

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
		
			// Format string content
			wchar_t buf[20];
			int len = swprintf(buf, 20, L"Test cell n° %d", i*j);
			tmp.bstrVal = SysAllocStringLen(buf, len);
			
			long indices[] = { i, j };
			SafeArrayPutElement(arr.parray, indices, (void *)&tmp);
		}
	}

	// Fill the workbook with theses data
	IDispatch *pXLSheet = sheet.GetActiveSheet(pXLApp);
	sheet.SetSheetName(pXLSheet, L"Data");
	IDispatch *pXLRange = sheet.GetRange(L"A2:O16", pXLSheet);
	sheet.SetValueInRange(arr, pXLRange);

	// Wait for user to see the result...
	MessageBox(NULL, L"All Done.", L"Notice", 0x10000);

	// Get value stored A4 cell
	IDispatch *pXLRange_2 = sheet.GetRange(L"A2:B4", pXLSheet);
	VARIANT cell = sheet.GetValue(pXLRange_2);
	if (cell.parray != NULL)
	{
		IDispatch *pXLSheets = sheet.GetAllSheets(pXLApp);
		IDispatch *pXLNewSheet = sheet.AddSheet(pXLSheets);
		sheet.SetSheetName(pXLNewSheet, L"Result");
		IDispatch *pXLRange_3 = sheet.GetRange(L"A1:B3", pXLNewSheet);
		sheet.SetValueInRange(cell, pXLRange_3);
		sheet.SetRangeColor(pXLRange_3, 255, 192, 0);

		MessageBox(NULL, L"All Done.", L"Notice", 0x10000);
	}

	// Close 
	// sheet.CloseInstance(pXLApp);

	pXLApp->Release();
	// pXLWorkbooks->Release();
	pXLBook->Release();
	pXLRange->Release();
	pXLRange_2->Release();
	pXLSheet->Release();
	VariantClear(&arr);
	VariantClear(&cell);

	// Unitialize COM
	CoUninitialize();
}