#include <stdio.h>
#include <comutil.h>
#include <string>
#include <time.h>
#include "stdafx.h"
#include "msoledbsql.h"
#include "dart_api.h"
#include "dart_native_api.h"
#include "dart_mssql.h"

BSTR intToBSTR(int number) {
	wchar_t buf[32];
	int len = swprintf(buf, 32, L"%d", number);
	BSTR result = SysAllocStringLen(buf, len);
	return result;
}

void addError(int* errorCount, std::string* errorMessages, std::string errorMessage) {
	(*errorMessages).append(errorMessage);
	(*errorCount)++;
}

BSTR UTF8ToBSTR(const char* utf8Value) {
	int wslen = MultiByteToWideChar(CP_UTF8, 0, utf8Value, (int)strlen(utf8Value), 0, 0);
	BSTR result = SysAllocStringLen(0, wslen);
	MultiByteToWideChar(CP_UTF8, 0, utf8Value, (int)strlen(utf8Value), result, wslen);
	return result;
}

char* WideCharToUTF8(const wchar_t* value) {
	if (value == NULL) {
		return NULL;
	}
	int n = WideCharToMultiByte(CP_UTF8, 0, value, -1, NULL, 0, NULL, NULL);
	if (n <= 0) {
		return NULL;
	}
	char* buffer = new char[n];
	WideCharToMultiByte(CP_UTF8, 0, value, -1, buffer, n, NULL, NULL);
	return buffer;
}

// DumpErrorInfo queries MSOLEDBSQL error interfaces, retrieving available  
// status or error information.  
void DumpErrorInfo(IUnknown* pObjectWithError, REFIID IID_InterfaceWithError, int* errorCount, std::string* errorMessages)
{

	// Interfaces used in the example.  
	IErrorInfo*             pIErrorInfoAll = NULL;
	IErrorInfo*             pIErrorInfoRecord = NULL;
	IErrorRecords*          pIErrorRecords = NULL;
	ISupportErrorInfo*      pISupportErrorInfo = NULL;
	ISQLErrorInfo*          pISQLErrorInfo = NULL;
	ISQLServerErrorInfo*    pISQLServerErrorInfo = NULL;

	// Number of error records.  
	ULONG                   nRecs;
	ULONG                   nRec;

	// Basic error information from GetBasicErrorInfo.  
	ERRORINFO               errorinfo;

	// IErrorInfo values.  
	BSTR                    bstrDescription;
	BSTR                    bstrSource;

	// ISQLErrorInfo parameters.  
	BSTR                    bstrSQLSTATE;
	LONG                    lNativeError;

	// ISQLServerErrorInfo parameter pointers.  
	SSERRORINFO*            pSSErrorInfo = NULL;
	OLECHAR*                pSSErrorStrings = NULL;

	// Hard-code an American English locale for the example.  
	DWORD                   MYLOCALEID = 0x0409;

	// printing current date and time (moacir):
	time_t					t = time(NULL);
	char					strTime[26];

	ctime_s(strTime, sizeof strTime, &t);
	fprintf_s(stderr, "Date/Time:\t%s", strTime);

	// Only ask for error information if the interface supports  
	// it.  
	if (FAILED(pObjectWithError->QueryInterface(IID_ISupportErrorInfo,
		(void**)&pISupportErrorInfo)))
	{
		fwprintf_s(stderr, L"SupportErrorErrorInfo interface not supported");
		return;
	}
	if (FAILED(pISupportErrorInfo->
		InterfaceSupportsErrorInfo(IID_InterfaceWithError)))
	{
		fwprintf_s(stderr, L"InterfaceWithError interface not supported");
		return;
	}

	// Do not test the return of GetErrorInfo. It can succeed and return  
	// a NULL pointer in pIErrorInfoAll. Simply test the pointer.  
	GetErrorInfo(0, &pIErrorInfoAll);

	if (pIErrorInfoAll != NULL)
	{
		// Test to see if it's a valid OLE DB IErrorInfo interface   
		// exposing a list of records.  
		if (SUCCEEDED(pIErrorInfoAll->QueryInterface(IID_IErrorRecords,
			(void**)&pIErrorRecords)))
		{
			pIErrorRecords->GetRecordCount(&nRecs);

			// Within each record, retrieve information from each  
			// of the defined interfaces.  
			for (nRec = 0; nRec < nRecs; nRec++)
			{
				// From IErrorRecords, get the HRESULT and a reference  
				// to the ISQLErrorInfo interface.  
				pIErrorRecords->GetBasicErrorInfo(nRec, &errorinfo);
				pIErrorRecords->GetCustomErrorObject(nRec,
					IID_ISQLErrorInfo, (IUnknown**)&pISQLErrorInfo);

				// Display the HRESULT, then use the ISQLErrorInfo.  
				fwprintf_s(stderr, L"HRESULT:\t%#X\n", errorinfo.hrError);

				if (pISQLErrorInfo != NULL)
				{
					pISQLErrorInfo->GetSQLInfo(&bstrSQLSTATE,
						&lNativeError);

					// Display the SQLSTATE and native error values.  
					fwprintf_s(stderr, L"SQLSTATE:\t%s\nNative Error:\t%ld\n", bstrSQLSTATE, lNativeError);

					// SysFree BSTR references.  
					SysFreeString(bstrSQLSTATE);

					// Get the ISQLServerErrorInfo interface from  
					// ISQLErrorInfo before releasing the reference.  
					pISQLErrorInfo->QueryInterface(
						IID_ISQLServerErrorInfo,
						(void**)&pISQLServerErrorInfo);

					pISQLErrorInfo->Release();
				}

				// Test to ensure the reference is valid, then  
				// get error information from ISQLServerErrorInfo.  
				if (pISQLServerErrorInfo != NULL)
				{
					pISQLServerErrorInfo->GetErrorInfo(&pSSErrorInfo,
						&pSSErrorStrings);

					// ISQLServerErrorInfo::GetErrorInfo succeeds  
					// even when it has nothing to return. Test the  
					// pointers before using.  
					if (pSSErrorInfo)
					{
						// Display the state and severity from the  
						// returned information. The error message comes  
						// from IErrorInfo::GetDescription.  
						fwprintf_s(stderr, L"Error state:\t%d\nSeverity:\t%d\n", pSSErrorInfo->bState, pSSErrorInfo->bClass);

						// IMalloc::Free needed to release references  
						// on returned values. For the example, assume  
						// the g_pIMalloc pointer is valid.  
						// Moacir: Troquei pra CoTaskMemFree pq nao achei lib de g_pIMalloc
						CoTaskMemFree(pSSErrorStrings);
						CoTaskMemFree(pSSErrorInfo);
					}

					pISQLServerErrorInfo->Release();
				}

				if (SUCCEEDED(pIErrorRecords->GetErrorInfo(nRec,
					MYLOCALEID, &pIErrorInfoRecord)))
				{
					// Get the source and description (error message)  
					// from the record's IErrorInfo.  
					pIErrorInfoRecord->GetSource(&bstrSource);
					pIErrorInfoRecord->GetDescription(&bstrDescription);

					if (bstrSource != NULL)
					{
						fwprintf_s(stderr, L"Source:\t\t%s\n", bstrSource);
						SysFreeString(bstrSource);
					}
					if (bstrDescription != NULL)
					{
						fwprintf_s(stderr, L"Error message:\t%s\n",	bstrDescription);
						SysFreeString(bstrDescription);
					}

					pIErrorInfoRecord->Release();
				}
			}

			pIErrorRecords->Release();
		}
		else
		{
			// IErrorInfo is valid; get the source and  
			// description to see what it is.  
			pIErrorInfoAll->GetSource(&bstrSource);
			pIErrorInfoAll->GetDescription(&bstrDescription);

			if (bstrSource != NULL)
			{
				fwprintf_s(stderr, L"Source:\t\t%s\n", bstrSource);
				SysFreeString(bstrSource);
			}
			if (bstrDescription != NULL)
			{
				fwprintf_s(stderr, L"Error message:\t%s\n", bstrDescription);
				SysFreeString(bstrDescription);
			}
		}

		pIErrorInfoAll->Release();
	}
	else
	{
		fwprintf_s(stderr, L"GetErrorInfo failed.");
	}

	pISupportErrorInfo->Release();

	return;
}

HRESULT oleCheck(HRESULT hr, IUnknown* pObjectWithError, REFIID IID_InterfaceWithError, int* errorCount, std::string* errorMessages) {
	if (FAILED(hr) && pObjectWithError != NULL) {
		DumpErrorInfo(pObjectWithError, IID_InterfaceWithError, errorCount, errorMessages);
	}
	// fill errorMessages if last operation had a error 
	if (FAILED(hr) && errorCount != NULL) {
		IErrorInfo		*pIErrorInfo = NULL;
		IErrorRecords	*pIErrorRecords;

		(*errorMessages).assign("");
		GetErrorInfo(0, &pIErrorInfo);
		if (pIErrorInfo != NULL && *errorCount == 0) {
			HRESULT hr1 = pIErrorInfo->QueryInterface(IID_IErrorRecords, (void **)&pIErrorRecords);
			if (SUCCEEDED(hr1)) {
				ULONG pcRecords;
				pIErrorRecords->GetRecordCount(&pcRecords);
				*errorCount = pcRecords;
				boolean ok = FALSE;
				for (ULONG i = 0; i < pcRecords; i++) {
					ISQLServerErrorInfo *pISQLServerErrorInfo;
					pIErrorRecords->GetCustomErrorObject(
						i, IID_ISQLServerErrorInfo,
						(IUnknown**)&pISQLServerErrorInfo);
					ok = (pISQLServerErrorInfo > 0);
					if (ok) {
						SSERRORINFO *errorInfo;
						OLECHAR *errorMsgs;
						pISQLServerErrorInfo->GetErrorInfo(&errorInfo, &errorMsgs);
						ok = (errorInfo != NULL);
						if (ok) {
							(*errorMessages).append(WideCharToUTF8(errorMsgs));
							CoTaskMemFree(errorInfo);
							CoTaskMemFree(errorMsgs);
						}
						pISQLServerErrorInfo->Release();
					}
					if (!ok) {
						OLECHAR *errorMsgs;
						pIErrorRecords->GetErrorInfo(i, 0, &pIErrorInfo);
						pIErrorInfo->GetDescription(&errorMsgs);
						ok = (errorMsgs != NULL);
						if (ok) {
							(*errorMessages).append(WideCharToUTF8(errorMsgs));
							pIErrorInfo->Release();
							SysFreeString(errorMsgs);
						}
					}
				}
				pIErrorRecords->Release();
			}
		}
		if ((*errorMessages).length() == 0) {
			char errorBuf[32];
			int len = snprintf(errorBuf, 32, "error code: 0x%X\0", hr);
			addError(errorCount, errorMessages, errorBuf);
		}
		if (pIErrorInfo != NULL) {
			pIErrorInfo->Release();
		}
	}
	return hr;
}

HRESULT memoryCheck(HRESULT hr, void *pv, int* errorCount, std::string* errorMessages) {
	if (!pv) {
		hr = E_OUTOFMEMORY;
		return oleCheck(hr, NULL, GUID_NULL, errorCount, errorMessages);
	}
	else {
		return hr;
	}
}

/////////////////////////////////////////////////////////////////
// freeBindings
//
//	This function frees a bindings array and any allocated
//	structures contained in that array.
//
/////////////////////////////////////////////////////////////////
void freeBindings(DBORDINAL	cBindings, DBBINDING* rgBindings) {
	ULONG iBind;

	// Free any memory used by DBOBJECT structures in the array
	for (iBind = 0; iBind < cBindings; iBind++)
		CoTaskMemFree(rgBindings[iBind].pObject);

	// Now free the bindings array itself
	CoTaskMemFree(rgBindings);
}

/////////////////////////////////////////////////////////////////
// getProperty
//
//	This function gets the BOOL value for the specified property
//	and returns the result in *pbValue.
//
//  here we do not pass errorCount and errorMessages since errors are expected for non-supported properties
/////////////////////////////////////////////////////////////////
HRESULT getProperty(IUnknown* pIUnknown, REFIID	riid, DBPROPID dwPropertyID, REFGUID guidPropertySet, BOOL* pbValue) {
	HRESULT					hr;
	DBPROPID				rgPropertyIDs[1];
	DBPROPIDSET				rgPropertyIDSets[1];

	ULONG					cPropSets = 0;
	DBPROPSET*				rgPropSets = NULL;

	IDBProperties*			pIDBProperties = NULL;
	ISessionProperties*		pISesProps = NULL;
	ICommandProperties*		pICmdProps = NULL;
	IRowsetInfo*			pIRowsetInfo = NULL;

	// Initialize the output value
	*pbValue = FALSE;

	// Set up the property ID array
	rgPropertyIDs[0] = dwPropertyID;

	// Set up the Property ID Set
	rgPropertyIDSets[0].rgPropertyIDs = rgPropertyIDs;
	rgPropertyIDSets[0].cPropertyIDs = 1;
	rgPropertyIDSets[0].guidPropertySet = guidPropertySet;

	// Get the property value for this property from the provider, but
	// don't try to display extended error information, since this may
	// not be a supported property: a failure is, in fact, expected if
	// the property is not supported
	if (riid == IID_IDBProperties) {
		hr = pIUnknown->QueryInterface(IID_IDBProperties,
			(void**)&pIDBProperties);
		if (oleCheck(hr, pIDBProperties, IID_IDBProperties, NULL, NULL) < 0) goto CLEANUP;
		hr = pIDBProperties->GetProperties(
			1,					//cPropertyIDSets
			rgPropertyIDSets,	//rgPropertyIDSets
			&cPropSets,			//pcPropSets
			&rgPropSets			//prgPropSets);
		);
		if (oleCheck(hr, pIDBProperties, IID_IDBProperties, NULL, NULL) < 0) goto CLEANUP;
	}
	else if (riid == IID_ISessionProperties) {
		hr = pIUnknown->QueryInterface(IID_ISessionProperties,
			(void**)&pISesProps);
		if (oleCheck(hr, pISesProps, IID_ISessionProperties, NULL, NULL) < 0) goto CLEANUP;
		hr = pISesProps->GetProperties(
			1,					//cPropertyIDSets
			rgPropertyIDSets,	//rgPropertyIDSets
			&cPropSets,			//pcPropSets
			&rgPropSets			//prgPropSets
		);
		if (oleCheck(hr, pISesProps, IID_ISessionProperties, NULL, NULL) < 0) goto CLEANUP;
	}
	else if (riid == IID_ICommandProperties) {
		hr = pIUnknown->QueryInterface(IID_ICommandProperties,
			(void**)&pICmdProps);
		if (oleCheck(hr, pICmdProps, IID_ICommandProperties, NULL, NULL) < 0) goto CLEANUP;
		hr = pICmdProps->GetProperties(
			1,					//cPropertyIDSets
			rgPropertyIDSets,	//rgPropertyIDSets
			&cPropSets,			//pcPropSets
			&rgPropSets			//prgPropSets
		);
		if (oleCheck(hr, pICmdProps, IID_ICommandProperties, NULL, NULL) < 0) goto CLEANUP;
	}
	else {
		hr = pIUnknown->QueryInterface(IID_IRowsetInfo,
			(void**)&pIRowsetInfo);
		if (oleCheck(hr, pIRowsetInfo, IID_IRowsetInfo, NULL, NULL) < 0) goto CLEANUP;
		hr = pIRowsetInfo->GetProperties(
			1,					//cPropertyIDSets
			rgPropertyIDSets,	//rgPropertyIDSets
			&cPropSets,			//pcPropSets
			&rgPropSets			//prgPropSets
		);
		if (oleCheck(hr, pIRowsetInfo, IID_IRowsetInfo, NULL, NULL) < 0) goto CLEANUP;
	}

	// Return the value for this property to the caller if
	// it's a VT_BOOL type value, as expected
	if (V_VT(&rgPropSets[0].rgProperties[0].vValue) == VT_BOOL)
		*pbValue = V_BOOL(&rgPropSets[0].rgProperties[0].vValue);

CLEANUP:
	if (rgPropSets) {
		CoTaskMemFree(rgPropSets[0].rgProperties);
		CoTaskMemFree(rgPropSets);
	}
	if (pIDBProperties)
		pIDBProperties->Release();
	if (pISesProps)
		pISesProps->Release();
	if (pICmdProps)
		pICmdProps->Release();
	if (pIRowsetInfo)
		pIRowsetInfo->Release();
	return hr;
}

HRESULT createShortDataAccessor(IUnknown* pUnkRowset, HACCESSOR* phAccessor, DBORDINAL* pcBindings, DBBINDING** prgBindings, DBORDINAL cColumns, DBCOLUMNINFO *rgColumnInfo, DBORDINAL* pcbRowSize, int* errorCount, std::string* errorMessages) {
	HRESULT					hr = NULL;
	IAccessor*				pIAccessor = NULL;
	DBLENGTH				iCol;
	DBORDINAL				dwOffset = 0;
	DBBINDING*				rgBindings = NULL;

	// Getting short data count
	int iShortCol = 0;
	for (iCol = 0; iCol < cColumns; iCol++) {
		// Only short Data!
		if (rgColumnInfo[iCol].wType != DBTYPE_IUNKNOWN && (rgColumnInfo[iCol].dwFlags & DBCOLUMNFLAGS_ISLONG) == 0)
			iShortCol++;
	}

	// Allocate memory for the bindings array; there is a one-to-one
	// mapping between the columns returned from GetColumnInfo and our
	// bindings
	if (iShortCol > 0) {
		rgBindings = (DBBINDING*)CoTaskMemAlloc(iShortCol * sizeof(DBBINDING));
		if (memoryCheck(hr, rgBindings, errorCount, errorMessages) < 0) goto CLEANUP;
		memset(rgBindings, 0, iShortCol * sizeof(DBBINDING));

		// Construct the binding array element for each column
		iShortCol = 0;
		for (iCol = 0; iCol < cColumns; iCol++) {
			if (rgColumnInfo[iCol].wType != DBTYPE_IUNKNOWN && (rgColumnInfo[iCol].dwFlags & DBCOLUMNFLAGS_ISLONG) == 0) {
				// This binding applies to the ordinal of this column
				rgBindings[iShortCol].iOrdinal = rgColumnInfo[iCol].iOrdinal;

				// We are asking the provider to give us the data for this column
				// (DBPART_VALUE), the length of that data (DBPART_LENGTH), and
				// the status of the column (DBPART_STATUS)
				rgBindings[iShortCol].dwPart = DBPART_VALUE | DBPART_LENGTH | DBPART_STATUS;

				// The following values are the offsets to the status, length, and
				// data value that the provider will fill with the appropriate values
				// when we fetch data later. When we fetch data, we will pass a
				// pointer to a buffer that the provider will copy column data to,
				// in accordance with the binding we have provided for that column;
				// these are offsets into that future buffer
				rgBindings[iShortCol].obStatus = dwOffset;
				rgBindings[iShortCol].obLength = dwOffset + sizeof(DBSTATUS);
				rgBindings[iShortCol].obValue = dwOffset + sizeof(DBSTATUS) + sizeof(DBORDINAL);
				rgBindings[iShortCol].pTypeInfo = NULL;
				rgBindings[iShortCol].pBindExt = NULL;
				rgBindings[iShortCol].pObject = NULL;
				rgBindings[iShortCol].dwFlags = 0;

				// Any memory allocated for the data value will be owned by us, the
				// client. Note that no data will be allocated in this case, as the
				// DBTYPE_WSTR bindings we are using will tell the provider to simply
				// copy data directly into our provided buffer
				rgBindings[iShortCol].dwMemOwner = DBMEMOWNER_CLIENTOWNED;

				// This is not a parameter binding
				rgBindings[iShortCol].eParamIO = DBPARAMIO_NOTPARAM;

				// We want to use the precision and scale of the column
				rgBindings[iShortCol].bPrecision = rgColumnInfo[iCol].bPrecision;
				rgBindings[iShortCol].bScale = rgColumnInfo[iCol].bScale;

				// Bind this column as DBTYPE_WSTR, which tells the provider to
				// copy a Unicode string representation of the data into our buffer,
				// converting from the native type if necessary
				rgBindings[iShortCol].wType = DBTYPE_WSTR;

				// Initially, we set the length for this data in our buffer to 0;
				// the correct value for this will be calculated directly below
				rgBindings[iShortCol].cbMaxLen = 0;

				// Determine the maximum number of bytes required in our buffer to
				// contain the Unicode string representation of the provider's native
				// data type, including room for the NULL-termination character
				switch (rgColumnInfo[iCol].wType)
				{
				case DBTYPE_NULL:
				case DBTYPE_EMPTY:
				case DBTYPE_I1:
				case DBTYPE_I2:
				case DBTYPE_I4:
				case DBTYPE_UI1:
				case DBTYPE_UI2:
				case DBTYPE_UI4:
				case DBTYPE_R4:
				case DBTYPE_BOOL:
				case DBTYPE_I8:
				case DBTYPE_UI8:
				case DBTYPE_R8:
				case DBTYPE_CY:
				case DBTYPE_ERROR:
					// When the above types are converted to a string, they
					// will all fit into 25 characters, so use that plus space
					// for the NULL-terminator
					rgBindings[iShortCol].cbMaxLen = (25 + 1) * sizeof(WCHAR);
					break;

				case DBTYPE_DECIMAL:
				case DBTYPE_NUMERIC:
				case DBTYPE_DATE:
				case DBTYPE_DBDATE:
				case DBTYPE_DBTIMESTAMP:
				case DBTYPE_GUID:
					// Converted to a string, the above types will all fit into
					// 50 characters, so use that plus space for the terminator
					rgBindings[iShortCol].cbMaxLen = (50 + 1) * sizeof(WCHAR);
					break;

				case DBTYPE_BYTES:
					// In converting DBTYPE_BYTES to a string, each byte
					// becomes two characters (e.g. 0xFF -> "FF"), so we
					// will use double the maximum size of the column plus
					// include space for the NULL-terminator
					rgBindings[iShortCol].cbMaxLen =
						(rgColumnInfo[iCol].ulColumnSize * 2 + 1) * sizeof(WCHAR);
					break;

				case DBTYPE_STR:
				case DBTYPE_WSTR:
				case DBTYPE_BSTR:
					// Going from a string to our string representation,
					// we can just take the maximum size of the column,
					// a count of characters, and include space for the
					// terminator, which is not included in the column size
					rgBindings[iShortCol].cbMaxLen =
						(rgColumnInfo[iCol].ulColumnSize + 1) * sizeof(WCHAR);
					break;

				default:
					// For any other type, we will simply use our maximum
					// column buffer size, since the display size of these
					// columns may be variable (e.g. DBTYPE_VARIANT) or
					// unknown (e.g. provider-specific types)
					if (*errorCount == 0) {
						addError(errorCount, errorMessages, "Unknown data type");
						hr = -1;
					}
					break;
				};

				// Ensure that the bound maximum length is no more than the
				// maximum column size in bytes that we've defined
				// (comentado por moacir. Prefiro pegar tudo!) rgBindings[iCol].cbMaxLen = min(rgBindings[iCol].cbMaxLen, maxColSize);

				// Update the offset past the end of this column's data, so
				// that the next column will begin in the correct place in
				// the buffer
				dwOffset = rgBindings[iShortCol].cbMaxLen + rgBindings[iShortCol].obValue;

				// Ensure that the data for the next column will be correctly
				// aligned for all platforms, or, if we're done with columns,
				// that if we allocate space for multiple rows that the data
				// for every row is correctly aligned
				dwOffset = ROUNDUP(dwOffset);
				iShortCol++;
			}
		}

		// Now that we have an array of bindings, tell the provider to
		// create the Accessor for those bindings. We get back a handle
		// to this Accessor, which we will use when fetching data
		hr = pUnkRowset->QueryInterface(
			IID_IAccessor, (void**)&pIAccessor);
		if (oleCheck(hr, pIAccessor, IID_IAccessor, errorCount, errorMessages) < 0) goto CLEANUP;
		hr = pIAccessor->CreateAccessor(
			DBACCESSOR_ROWDATA,	//dwAccessorFlags
			iShortCol,			//cBindings
			rgBindings,			//rgBindings
			0,					//cbRowSize
			phAccessor,			//phAccessor	
			NULL				//rgStatus
		);
		if (oleCheck(hr, pIAccessor, IID_IAccessor, errorCount, errorMessages) < 0) goto CLEANUP;
	}

	// Return the row size (the current dwOffset is the size of the row),
	// the count of bindings, and the bindings array to the caller
	*pcbRowSize = dwOffset;
	*pcBindings = iShortCol;
	*prgBindings = rgBindings;

CLEANUP:
	if (pIAccessor)
		pIAccessor->Release();
	return hr;
}

HRESULT createLargeDataAcessors(IUnknown* pUnkRowset, HACCESSOR** phAccessors, DBORDINAL* pcBindings, DBBINDING** prgBindings, DBORDINAL cColumns, DBCOLUMNINFO *rgColumnInfo, int* errorCount, std::string* errorMessages) {
	HRESULT					hr = NULL;
	DBLENGTH				iCol;
	DBBINDING*				rgBindings = NULL;
	IAccessor*				pIAccessor = NULL;
	HACCESSOR				hAccessor;
	ULONG					ulbindstatus;
	int						iLargeCol = 0;


	*phAccessors = NULL;
	
	// Getting large data count
	for (iCol = 0; iCol < cColumns; iCol++) {
		// Only short Data!
		if (rgColumnInfo[iCol].wType == DBTYPE_IUNKNOWN || (rgColumnInfo[iCol].dwFlags & DBCOLUMNFLAGS_ISLONG) != 0) 
			iLargeCol++;
	}

	// Allocate memory for the bindings array; there is a one-to-one
	// mapping between the columns returned from GetColumnInfo and our
	// bindings
	if (iLargeCol > 0) {
		rgBindings = (DBBINDING*)CoTaskMemAlloc(iLargeCol * sizeof(DBBINDING));
		if (memoryCheck(hr, rgBindings, errorCount, errorMessages) < 0) goto CLEANUP;
		memset(rgBindings, 0, iLargeCol * sizeof(DBBINDING));

		// Construct the binding array element for each column
		iLargeCol = 0;
		for (iCol = 0; iCol < cColumns; iCol++) {
			if (rgColumnInfo[iCol].wType == DBTYPE_IUNKNOWN || (rgColumnInfo[iCol].dwFlags & DBCOLUMNFLAGS_ISLONG) != 0) {
				// This binding applies to the ordinal of this column
				rgBindings[iLargeCol].iOrdinal = rgColumnInfo[iCol].iOrdinal;

				// We are asking the provider to give us the data for this column
				// (DBPART_VALUE) and the status of the column (DBPART_STATUS)
				rgBindings[iLargeCol].dwPart = DBPART_VALUE | DBPART_STATUS;

				// The following values are the offsets to the status, length, and
				// data value that the provider will fill with the appropriate values
				// when we fetch data later. When we fetch data, we will pass a
				// pointer to a buffer that the provider will copy column data to,
				// in accordance with the binding we have provided for that column;
				// these are offsets into that future buffer
				rgBindings[iLargeCol].obStatus = sizeof(IUnknown*);
				rgBindings[iLargeCol].obLength = 0;
				rgBindings[iLargeCol].obValue = 0;
				rgBindings[iLargeCol].pTypeInfo = NULL;
				rgBindings[iLargeCol].pBindExt = NULL;
				rgBindings[iLargeCol].dwFlags = 0;
				rgBindings[iLargeCol].bPrecision = 0;
				rgBindings[iLargeCol].bScale = 0;
				// To specify the type of object that we want from the
				// provider, we need to create a DBOBJECT structure and
				// place it in our binding for this column
				rgBindings[iLargeCol].pObject =
					(DBOBJECT *)CoTaskMemAlloc(sizeof(DBOBJECT));
				if (memoryCheck(hr, rgBindings[iLargeCol].pObject, errorCount, errorMessages) < 0) goto CLEANUP;
				// Direct the provider to create an ISequentialStream
				// object over the data for this column
				rgBindings[iLargeCol].pObject->iid = IID_ISequentialStream;
				// We want read access on the ISequentialStream
				// object that the provider will create for us
				rgBindings[iLargeCol].pObject->dwFlags = STGM_READ;

				// Any memory allocated for the data value will be owned by us, the
				// client. Note that no data will be allocated in this case, as the
				// DBTYPE_WSTR bindings we are using will tell the provider to simply
				// copy data directly into our provided buffer
				rgBindings[iLargeCol].dwMemOwner = DBMEMOWNER_CLIENTOWNED;

				// This is not a parameter binding
				rgBindings[iLargeCol].eParamIO = DBPARAMIO_NOTPARAM;

				// To create an ISequentialStream object, we will
				// bind this column as DBTYPE_IUNKNOWN to indicate
				// that we are requesting this column as an object
				rgBindings[iLargeCol].wType = DBTYPE_IUNKNOWN;

				// We want to allocate enough space in our buffer for
				// the ISequentialStream pointer we will obtain from
				// the provider
				rgBindings[iLargeCol].cbMaxLen = 0;

				// Ensure that the bound maximum length is no more than the
				// maximum column size in bytes that we've defined
				// (comentado por moacir. Prefiro pegar tudo!) rgBindings[iCol].cbMaxLen = min(rgBindings[iCol].cbMaxLen, maxColSize);

				// Now that we have an one bindings for lage object, tell the provider to
				// create the Accessor for this bindins. We get back a handle
				// to this Accessor, which we will use when fetching data
				hr = pUnkRowset->QueryInterface(
					IID_IAccessor, (void**)&pIAccessor);
				if (oleCheck(hr, pIAccessor, IID_IAccessor, errorCount, errorMessages) < 0) goto CLEANUP;
				hr = pIAccessor->CreateAccessor(
					DBACCESSOR_ROWDATA,		//dwAccessorFlags
					1,						//cBindings
					&rgBindings[iLargeCol],	//rgBindings
					0,						//cbRowSize
					&hAccessor,				//phAccessor	
					&ulbindstatus    		//rgStatus
				);
				*phAccessors = (HACCESSOR*)CoTaskMemRealloc(*phAccessors, (iLargeCol + 1) * sizeof(HACCESSOR));
				if (memoryCheck(hr, *phAccessors, errorCount, errorMessages) < 0) goto CLEANUP;

				(*phAccessors)[iLargeCol] = hAccessor;
				if (pIAccessor)
					pIAccessor->Release();

				iLargeCol++;
			}
		}
	}
	*pcBindings = iLargeCol;
	*prgBindings = rgBindings;

CLEANUP:
	return hr;
}

HRESULT createParamsAccessor(IUnknown* pUnkRowset, HACCESSOR* phAccessor, DBCOUNTITEM sqlParamsCount, Dart_Handle sqlParams, DBORDINAL* pcbRowSize, void** ppParamData, int* errorCount, std::string* errorMessages) {
	HRESULT					hr = NULL;
	IAccessor*				pIAccessor = NULL;
	DBCOUNTITEM				iCol;
	DBORDINAL				dwOffset = 0, lastOffset = 0;
	DBBINDING*				rgBindings = NULL;
	void*					pParamData = NULL;
	void*					pCurParam = NULL;
	void*					pSourceValue = NULL;

	// Allocate memory for the bindings array; there is a one-to-one
	// mapping between the columns returned from GetColumnInfo and our
	// bindings
	rgBindings = (DBBINDING*)CoTaskMemAlloc(sqlParamsCount * sizeof(DBBINDING));
	if (memoryCheck(hr, rgBindings, errorCount, errorMessages) < 0) goto CLEANUP;
	memset(rgBindings, 0, sqlParamsCount * sizeof(DBBINDING));

	// Construct the binding array element for each param
	for (iCol = 0; iCol < sqlParamsCount; iCol++) {
		int64_t dartValueI64;
		const char* dartValueString;
		// This binding applies to the ordinal of this column
		rgBindings[iCol].iOrdinal = iCol+1;
		rgBindings[iCol].dwPart = DBPART_VALUE;
		rgBindings[iCol].obStatus = 0; 
		rgBindings[iCol].obLength = 0; 
		rgBindings[iCol].obValue = dwOffset;
		rgBindings[iCol].pTypeInfo = NULL;
		rgBindings[iCol].pBindExt = NULL;
		rgBindings[iCol].pObject = NULL;
		rgBindings[iCol].dwFlags = 0;
		rgBindings[iCol].dwMemOwner = DBMEMOWNER_CLIENTOWNED;
		rgBindings[iCol].eParamIO = DBPARAMIO_INPUT;
		rgBindings[iCol].bPrecision = 0;
		rgBindings[iCol].bScale = 0;
		Dart_Handle item = Dart_Handle(Dart_ListGetAt(sqlParams, iCol));
		if (Dart_IsNumber(item)) {
			HandleError(Dart_IntegerToInt64(item, &dartValueI64));
			rgBindings[iCol].wType = DBTYPE_I4;
			rgBindings[iCol].cbMaxLen = sizeof(dartValueI64);
			pSourceValue = &dartValueI64;
		}
		else if (Dart_IsString(item)) {
			HandleError(Dart_StringToCString(item, &dartValueString));			
			rgBindings[iCol].wType = DBTYPE_STR;
			rgBindings[iCol].cbMaxLen = strlen(dartValueString);
			pSourceValue = (void*) dartValueString;
		}
		else {
			addError(errorCount, errorMessages, "Unknown param type");
		}
		
		lastOffset = dwOffset;
		dwOffset = dwOffset + rgBindings[iCol].cbMaxLen;
		dwOffset = ROUNDUP(dwOffset);

		pParamData = CoTaskMemRealloc(pParamData, dwOffset);
		if (memoryCheck(hr, pParamData, errorCount, errorMessages) < 0) goto CLEANUP;

		void* pCurParam = (BYTE*)pParamData + lastOffset;
		memset(pCurParam, 0, dwOffset - lastOffset);
		memcpy((void*)pCurParam, pSourceValue, rgBindings[iCol].cbMaxLen);
	}

	// Now that we have an array of bindings, tell the provider to
	// create the Accessor for those bindings. We get back a handle
	// to this Accessor, which we will use when fetching data
	hr = pUnkRowset->QueryInterface(
		IID_IAccessor, (void**)&pIAccessor);
	if (oleCheck(hr, pIAccessor, IID_IAccessor, errorCount, errorMessages) < 0) goto CLEANUP;
	hr = pIAccessor->CreateAccessor(
		DBACCESSOR_PARAMETERDATA,	//dwAccessorFlags
		sqlParamsCount,				//cBindings
		rgBindings,					//rgBindings
		dwOffset,					//cbRowSize
		phAccessor,					//phAccessor	
		NULL						//rgStatus
	);
	if (oleCheck(hr, pIAccessor, IID_IAccessor, errorCount, errorMessages) < 0) goto CLEANUP;

	*pcbRowSize = dwOffset;
	*ppParamData = pParamData;

CLEANUP:
	if (pIAccessor)
		pIAccessor->Release();
	return hr;
}

