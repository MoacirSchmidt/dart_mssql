#include <atlbase.h>
#include <string>
#include "msoledbsql.h"
#include "stdafx.h"
#include "oledberr.h"
#include "dart_api.h"
#include "dart_native_api.h"
#include "dart_mssql.h"

#define AUTH_TYPE_PASWORD		0
#define AUTH_TYPE_INTEGRATED	1

#define MAX_DISPLAY_SIZE		100

using namespace std;

HRESULT sqlConnect(const char* serverName, const char* dbName, const char* userName, const char* password, int64_t authType, IDBInitialize** ppInitialize, int* errorCount, std::string* errorMessages) {
	IDBInitialize*		pIDBInitialize = nullptr;
	IDBProperties*		pIDBProperties = nullptr;
	DBPROP				InitProperties[4] = { 0 };
	DBPROPSET			rgInitPropSet[1] = { 0 };
	HRESULT				hr = S_OK;
	int					numProperties;

	*ppInitialize = NULL;
	
	// Obtain access to the OLE DB Driver for SQL Server.  
	hr = CoCreateInstance(CLSID_MSOLEDBSQL,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_IDBInitialize,
		(void **)&pIDBInitialize);
	if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;

	// Initialize property values needed to establish connection.  
	for (int i = 0; i < 4; i++) {
		VariantInit(&InitProperties[i].vValue);
	}

	// Server name.  
	// See DBPROP structure for more information on InitProperties  
	InitProperties[0].dwPropertyID = DBPROP_INIT_DATASOURCE;
	InitProperties[0].vValue.vt = VT_BSTR;
	InitProperties[0].vValue.bstrVal = UTF8ToBSTR(serverName);
	InitProperties[0].dwOptions = DBPROPOPTIONS_REQUIRED;
	InitProperties[0].colid = DB_NULLID;

	// Database.  
	InitProperties[1].dwPropertyID = DBPROP_INIT_CATALOG;
	InitProperties[1].vValue.vt = VT_BSTR;
	InitProperties[1].vValue.bstrVal = UTF8ToBSTR(dbName);
	InitProperties[1].dwOptions = DBPROPOPTIONS_REQUIRED;
	InitProperties[1].colid = DB_NULLID;

	// Username (login).  
	if (authType == AUTH_TYPE_INTEGRATED) {
		InitProperties[2].dwPropertyID = DBPROP_AUTH_INTEGRATED;
		InitProperties[2].vValue.vt = VT_BSTR;
		InitProperties[2].vValue.bstrVal = SysAllocString(L"SSPI");
		InitProperties[2].dwOptions = DBPROPOPTIONS_REQUIRED;
		InitProperties[2].colid = DB_NULLID;
		numProperties = 3;
	}
	else {
		InitProperties[2].dwPropertyID = DBPROP_AUTH_USERID;
		InitProperties[2].vValue.vt = VT_BSTR;
		InitProperties[2].vValue.bstrVal = UTF8ToBSTR(userName);
		InitProperties[2].dwOptions = DBPROPOPTIONS_REQUIRED;
		InitProperties[2].colid = DB_NULLID;

		// password
		InitProperties[3].dwPropertyID = DBPROP_AUTH_PASSWORD;
		InitProperties[3].vValue.vt = VT_BSTR;
		InitProperties[3].vValue.bstrVal = UTF8ToBSTR(password);
		InitProperties[3].dwOptions = DBPROPOPTIONS_REQUIRED;
		InitProperties[3].colid = DB_NULLID;
		numProperties = 4;
	}

	// Construct the DBPROPSET structure(rgInitPropSet). The   
	// DBPROPSET structure is used to pass an array of DBPROP   
	// structures (InitProperties) to the SetProperties method.  
	rgInitPropSet[0].guidPropertySet = DBPROPSET_DBINIT;
	rgInitPropSet[0].cProperties = numProperties;
	rgInitPropSet[0].rgProperties = InitProperties;

	// Set initialization properties.  
	hr = pIDBInitialize->QueryInterface(IID_IDBProperties, (void **)&pIDBProperties);
	if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;

	hr = pIDBProperties->SetProperties(1, rgInitPropSet);
	if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;

	// Now establish the connection to the data source.  
	hr = pIDBInitialize->Initialize();
	if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;

	// I'm not sure if we need that
	for (int i = 0; i < 4; i++) {
		SysFreeString(InitProperties[i].vValue.bstrVal);
	}

	*ppInitialize = pIDBInitialize;

CLEANUP:
	if (pIDBProperties)	{
		pIDBProperties->Release();
		pIDBProperties = nullptr;
	}
	return hr;
}

Dart_Handle dbColumnValueToDartHandle(DBBINDING* rgBindings, DBCOLUMNINFO *rgColumnInfo, void* pData, ULONG iBindingCol, ULONG iInfoCol, int* errorCount, std::string* errorMessages) {
	DBSTATUS				dwStatus;
	ULONG					ulLength;
	void*					pvValue;
	ISequentialStream*		pISeqStream = NULL;
	ULONG					cbRead;
	char*					c = NULL;
	char*					longValueBuf = NULL;
	HRESULT					hr;
	BYTE					pbBuff[1000];
	ULONG					cbNeeded = sizeof(pbBuff) / sizeof(BYTE);
	ULONG					cbReadTotal = 0;

	// We have bound status, length, and the data value for all
	// columns, so we know that these can all be used
	dwStatus = *(DBSTATUS *)((BYTE *)pData + rgBindings[iBindingCol].obStatus);
	ulLength = *(ULONG *)((BYTE *)pData + rgBindings[iBindingCol].obLength);
	pvValue = (BYTE *)pData + rgBindings[iBindingCol].obValue;
	Dart_Handle result = Dart_Null();

	// Check the status of this column. This decides
	// exactly what will be displayed for the column
	switch (dwStatus) {
	// The data is NULL, so don't try to display it
	case DBSTATUS_S_ISNULL:
		result = Dart_Null();
		break;

	// The data was fetched, but may have been truncated.
	// Display string data for this column to the user
	case DBSTATUS_S_TRUNCATED:
	case DBSTATUS_S_OK:
	case DBSTATUS_S_DEFAULT:
		// We have either bound the column as a Unicode string
		// (DBTYPE_WSTR) or as an ISequentialStream object
		// (DBTYPE_IUNKNOWN), and have to do different processing
		// for each one of these possibilities
		switch (rgBindings[iBindingCol].wType) {
		case DBTYPE_WSTR:
			switch (rgColumnInfo[iInfoCol].wType) {
			case DBTYPE_NUMERIC:
			case DBTYPE_I4:
				result = HandleError(Dart_NewInteger(wcstol((wchar_t *)pvValue, NULL, 10)));
				break;
			default:
				if (ulLength == 0) { // caso especial de string vazia ('')
					result = Dart_EmptyString();
				}
				else {
					// Copy the string data
					c = new char[ulLength * 2];
					WideCharToMultiByte(CP_UTF8, 0, (wchar_t *)pvValue, -1, c, ulLength * 2, NULL, NULL);
					result = HandleError(Dart_NewStringFromCString(c));
					delete[] c;
				}
			}
			break;
			case DBTYPE_IUNKNOWN:
				// We've bound this as an ISequentialStream object,
				// therefore the data in our buffer is a pointer
				// to the object's ISequentialStream interface

				pISeqStream = *(ISequentialStream**)pvValue;

				// We call ISequentialStream::Read to read bytes from
				// the stream blindly into our buffer, simply as a
				// demonstration of ISequentialStream. To display the
				// data properly, the native provider type of this
				// column should be accounted for; it could be
				// DBTYPE_WSTR, in which case this works, or it could
				// be DBTYPE_STR or DBTYPE_BYTES, in which case this
				// won't display the data correctly

				// Isto aqui está errado pois nao está convertendo pra UTF8. 
				longValueBuf = NULL;
				do {
					hr = pISeqStream->Read(pbBuff, cbNeeded, &cbRead);
					longValueBuf = (char *)CoTaskMemRealloc(longValueBuf, cbReadTotal + cbRead + 2);
					if (hr != S_FALSE) {
						memcpy((void*)&longValueBuf[cbReadTotal], pbBuff, cbRead);
						cbReadTotal += cbRead;
					}
				}
				while (SUCCEEDED(hr) && hr != S_FALSE && cbRead == cbNeeded);

				if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;

				// Since streams don't provide NULL-termination,
				// we'll NULL-terminate the resulting string ourselves
				longValueBuf[cbReadTotal] = '\0'; longValueBuf[cbReadTotal+1] = '\0';

				if (cbReadTotal == 0) { // caso especial de string vazia ('')
					result = Dart_EmptyString();
				}
				else {
					c = new char[cbReadTotal * 2];
					WideCharToMultiByte(CP_UTF8, 0, (wchar_t *)longValueBuf, -1, c, cbReadTotal * 2, NULL, NULL);
					result = HandleError(Dart_NewStringFromCString(c));
					delete[] c;
				}
				CoTaskMemFree(longValueBuf);

				// Release the stream object, now that we're done
				pISeqStream->Release();
				pISeqStream = NULL;
			break;
		default:
			result = Dart_Null();
			addError(errorCount, errorMessages, "Unknown column type");
			break;
		}
	}
CLEANUP:
	if (pISeqStream) {
		pISeqStream->Release();
	}
	return result;
}

void sqlExecute(IDBInitialize* pInitialize, const LPCOLESTR sqlCommand, Dart_Handle sqlParams, Dart_Handle dartMssqlLib, Dart_Handle dartSqlResult, int* errorCount, std::string* errorMessages) {
	IUnknown*			pSession = NULL; 
	IDBCreateSession*	pIDBCreateSession = NULL;
	IDBCreateCommand*	pIDBCreateCommand = NULL;
	ICommandText*		pICommandText = NULL;
	IColumnsInfo*		pIColumnsInfo = NULL;
	IRowset*			pIRowset = NULL;
	HRESULT				hr = S_OK;
	DBORDINAL			cColumns;
	DBORDINAL			cShortBindings = 0;
	DBBINDING*			rgShortBindings = NULL;
	DBORDINAL			cLargeBindings = 0;
	DBBINDING*			rgLargeBindings = NULL;
	DBCOLUMNINFO*       rgColumnInfo = NULL;
	LPWSTR				pStringBuffer = NULL;
	void*				pShortData = NULL;
	void*				pLargeData = NULL;
	void*				pParamData = NULL;
	DBCOUNTITEM			iRow = 0;
	DBROWCOUNT			cRowsAffected;
	HACCESSOR*			hLargeAccessors = NULL;
	HACCESSOR			hParamsAccessor = DB_NULL_HACCESSOR;
	DBPARAMS*			pParams;
	DBORDINAL			cbParamsRowSize;
	intptr_t			sqlParamsCount;
	Dart_Handle			sqlRowClass = HandleError(Dart_GetType(dartMssqlLib, Dart_NewStringFromCString("SqlRow"), 0, NULL));
	Dart_Handle			dartSqlResultColumns = NULL;
	Dart_Handle			dartSqlResultRows = NULL;
	Dart_Handle*		dartRows = NULL;

	hr = pInitialize->QueryInterface(IID_IDBCreateSession, (void**)&pIDBCreateSession);
	if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;

	hr = pIDBCreateSession->CreateSession(NULL, IID_IDBCreateCommand, &pSession);
	if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;

	hr = pSession->QueryInterface(IID_IDBCreateCommand, (void**)&pIDBCreateCommand);
	if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;

	hr = pIDBCreateCommand->CreateCommand(NULL,	IID_ICommandText, (IUnknown**)&pICommandText);
	if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;

	hr = pICommandText->SetCommandText(DBGUID_DBSQL, sqlCommand);
	if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;

	if (sqlParams == NULL || Dart_IsNull(sqlParams))
		sqlParamsCount = 0;
	else
		HandleError(Dart_ListLength(sqlParams, &sqlParamsCount));

	if (sqlParamsCount != 0) {
		hr = createParamsAccessor(pICommandText, &hParamsAccessor, sqlParamsCount, sqlParams, &cbParamsRowSize, &pParamData, errorCount, errorMessages);
		if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;
		pParams = new DBPARAMS;
		pParams->cParamSets = 1;
		pParams->hAccessor = hParamsAccessor;
		pParams->pData = pParamData;
	}
	else {
		pParams = NULL;
	}		

	hr = pICommandText->Execute(NULL, IID_IRowset, pParams, &cRowsAffected, (IUnknown**)&pIRowset);
	if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;

	if (pIRowset != NULL) {
		HACCESSOR	hShortAccessor = DB_NULL_HACCESSOR;
		DBORDINAL	cbShortRowSize;
		DBCOUNTITEM cRowsReturned, totalRows = 0;
		HROW*		pRow = NULL;
		Dart_Handle dartValuesName = Dart_NewStringFromCString("_values");
		Dart_Handle dartSqlResultName = Dart_NewStringFromCString("_sqlResult");
		
		// Obtain the column information for the Rowset; from this, we can find
		hr = pIRowset->QueryInterface(
			IID_IColumnsInfo, (void**)&pIColumnsInfo);
		if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;
		hr = pIColumnsInfo->GetColumnInfo(
			&cColumns,		//pcColumns
			&rgColumnInfo,	//prgColumnInfo
			&pStringBuffer	//ppStringBuffer
		);
		if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;

		// Create acessor for short data columns
		hr = createShortDataAccessor(pIRowset, &hShortAccessor, &cShortBindings, &rgShortBindings, cColumns, rgColumnInfo, &cbShortRowSize, errorCount, errorMessages);
		if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;

		// Create acessor for large data columns
		hr = createLargeDataAcessors(pIRowset, &hLargeAccessors, &cLargeBindings, &rgLargeBindings, cColumns, rgColumnInfo, errorCount, errorMessages);

		dartSqlResultColumns = HandleError(Dart_NewListOf(Dart_CoreType_String, cColumns));
		for (DBORDINAL iCol = 0; iCol < cColumns; iCol++) {
			char* buffer;
			buffer = WideCharToUTF8(rgColumnInfo[iCol].pwszName);
			Dart_Handle colName = Dart_NewStringFromCString(buffer);
			HandleError(Dart_ListSetAt(dartSqlResultColumns, iCol, colName));
			delete[] buffer;
		}

		// Allocate enough memory to hold a rows of data; this is
		// where the actual row data from the provider will be placed
		if (cShortBindings > 0) {
			pShortData = CoTaskMemAlloc(cbShortRowSize);
			if (memoryCheck(hr, pShortData, errorCount, errorMessages) < 0) goto CLEANUP;
		}

		if (cLargeBindings > 0) {
			pLargeData = CoTaskMemAlloc(sizeof(ISequentialStream *) + sizeof(ULONG));
			if (memoryCheck(hr, pLargeData, errorCount, errorMessages) < 0) goto CLEANUP;
		}

		while (hr == S_OK) {
			Dart_Handle	dartSqlRow;
			Dart_Handle	dartSqlRowValues;
			ULONG		iShortCol = 0;
			ULONG		iLargeCol = 0;

			// Attempt to get row handle from the provider
			hr = pIRowset->GetNextRows(DB_NULL_HCHAPTER, 0, 1, &cRowsReturned, &pRow);
			if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;

			if (cRowsReturned == 1) {
				dartSqlRow = HandleError(Dart_New(sqlRowClass, Dart_Null(), 0, NULL));
				dartSqlRowValues = HandleError(Dart_NewList(cColumns));

				// get short pData;
				if (cShortBindings > 0) {
					hr = pIRowset->GetData(pRow[0], hShortAccessor, pShortData);
					if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;
				}

				for (ULONG iCol = 0; iCol < cColumns; iCol++) {
					Dart_Handle colValue;
					if (rgColumnInfo[iCol].wType == DBTYPE_IUNKNOWN || (rgColumnInfo[iCol].dwFlags & DBCOLUMNFLAGS_ISLONG) != 0) {
						hr = pIRowset->GetData(pRow[0], hLargeAccessors[iLargeCol], pLargeData);
						if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;
						colValue = dbColumnValueToDartHandle(rgLargeBindings, rgColumnInfo, pLargeData, iLargeCol, iCol, errorCount, errorMessages);
						iLargeCol++;
					}
					else {
						colValue = dbColumnValueToDartHandle(rgShortBindings, rgColumnInfo, pShortData, iShortCol, iCol, errorCount, errorMessages);
						iShortCol++;
					}
					HandleError(Dart_ListSetAt(dartSqlRowValues, iCol, colValue)); 
				}
				HandleError(Dart_SetField(dartSqlRow, dartValuesName, dartSqlRowValues));
				HandleError(Dart_SetField(dartSqlRow, dartSqlResultName, dartSqlResult));

				hr = pIRowset->ReleaseRows(cRowsReturned, pRow, NULL, NULL, NULL);
				if (oleCheck(hr, errorCount, errorMessages) < 0) goto CLEANUP;
				CoTaskMemFree(pRow);
				pRow = NULL;

				totalRows += cRowsReturned;
				dartRows = (Dart_Handle*)CoTaskMemRealloc(dartRows, totalRows * sizeof(Dart_Handle*));
				dartRows[totalRows - 1] = dartSqlRow;
			}
		}
		cRowsAffected = totalRows;
		dartSqlResultRows = HandleError(Dart_NewList(totalRows));
		for (DBLENGTH i = 0; i < totalRows; i++) {
			HandleError(Dart_ListSetAt(dartSqlResultRows, i, dartRows[i]));
		}
		CoTaskMemFree(dartRows);
		HandleError(Dart_SetField(dartSqlResult, Dart_NewStringFromCString("rows"), dartSqlResultRows));
		HandleError(Dart_SetField(dartSqlResult, Dart_NewStringFromCString("columns"), dartSqlResultColumns));
	}

	HandleError(Dart_SetField(dartSqlResult, Dart_NewStringFromCString("rowsAffected"), HandleError(Dart_NewInteger(cRowsAffected))));

CLEANUP:
	if (pIRowset != NULL) {
		pIRowset->Release();
	}
	if (pICommandText != NULL) {
		pICommandText->Release();
	}
	if (pIDBCreateCommand != NULL) {
		pIDBCreateCommand->Release();
	}
	if (pSession != NULL) {
		pSession->Release();
	}
	if (pIDBCreateSession != NULL) {
		pIDBCreateSession->Release();
	}
	if (pIColumnsInfo != NULL)
		pIColumnsInfo->Release();

	CoTaskMemFree(pStringBuffer);
	CoTaskMemFree(rgColumnInfo);
	freeBindings(cShortBindings, rgShortBindings);
	freeBindings(cLargeBindings, rgLargeBindings);
	CoTaskMemFree(pShortData);
	CoTaskMemFree(pLargeData);
	CoTaskMemFree(pParamData);
	CoTaskMemFree(hLargeAccessors);
};

