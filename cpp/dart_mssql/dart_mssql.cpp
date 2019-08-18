// dart_mssql.cpp : Defines the exported functions for the DLL application.
//
#include <comdef.h>     
#include <comutil.h>
#include <map>
#include <exception>
#include "msoledbsql.h"
#include "dart_api.h"
#include "dart_native_api.h"
#include "dart_mssql.h"

Dart_Handle HandleError(Dart_Handle handle) {
	if (Dart_IsError(handle)) Dart_PropagateError(handle);
	return handle;
}

Dart_Handle	dartMssqlLib;
Dart_Handle	sqlResultClass;
Dart_Handle	sqlReturnClass;

void registerHandleInGlobalInterfaceTable(IDBInitialize* pSession, Dart_Handle dartSqlReturn, int* errorCount, std::string* errorMessages) {
	HRESULT hr;

	IGlobalInterfaceTable* pGIT = NULL;
	hr = CoCreateInstance(CLSID_StdGlobalInterfaceTable,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_IGlobalInterfaceTable,
		(void**)&pGIT);
	DWORD dwCookie = 0;
	if (oleCheck(hr, pSession, IID_IDBInitialize, errorCount, errorMessages) >= 0) {
		hr = pGIT->RegisterInterfaceInGlobal(pSession, IID_IDBInitialize, &dwCookie);
		if (oleCheck(hr, pSession, IID_IDBInitialize, errorCount, errorMessages) >= 0) {
			HandleError(Dart_SetField(dartSqlReturn, Dart_NewStringFromCString("handle"), Dart_NewIntegerFromUint64(dwCookie)));
		}
	}
}

IDBInitialize* getHandleFromGlobalInterfaceTable(DWORD cookie, int* errorCount, std::string* errorMessages) {
	IDBInitialize* pSession = NULL;
	HRESULT hr;
	IGlobalInterfaceTable* pGIT = NULL;
	hr = CoCreateInstance(CLSID_StdGlobalInterfaceTable,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_IGlobalInterfaceTable,
		(void**)&pGIT);
	if (oleCheck(hr, NULL, GUID_NULL, errorCount, errorMessages) >= 0) {
		hr = pGIT->GetInterfaceFromGlobal(cookie, IID_IDBInitialize, (void**)&pSession);
		oleCheck(hr, NULL, GUID_NULL, errorCount, errorMessages);
	}
	return pSession;
}

void revokeInterfaceFromGlobalInterfaceTable(DWORD cookie) {
	HRESULT hr;
	IGlobalInterfaceTable* pGIT = NULL;
	hr = CoCreateInstance(CLSID_StdGlobalInterfaceTable,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_IGlobalInterfaceTable,
		(void**)&pGIT);
	if (hr == S_OK) {
		hr = pGIT->RevokeInterfaceFromGlobal(cookie);
	}
}

void _connectCommand(Dart_NativeArguments arguments) {
	const char *serverName;
	const char *databaseName;
	const char *userName;
	const char *password;
	int64_t authType;
	int	 errorCount = 0;
	std::string errorMessages;
	IDBInitialize* pSession = NULL;
	HRESULT		   hr = S_OK;

	Dart_EnterScope();
	
	HandleError(Dart_StringToCString(HandleError(Dart_GetNativeArgument(arguments, 0)), &serverName));
	HandleError(Dart_StringToCString(HandleError(Dart_GetNativeArgument(arguments, 1)), &databaseName));
	HandleError(Dart_StringToCString(HandleError(Dart_GetNativeArgument(arguments, 2)), &userName));
	HandleError(Dart_StringToCString(HandleError(Dart_GetNativeArgument(arguments, 3)), &password));
	HandleError(Dart_IntegerToInt64(HandleError(Dart_GetNativeArgument(arguments, 4)), &authType));
	dartMssqlLib = HandleError(Dart_LookupLibrary(Dart_NewStringFromCString("package:dart_mssql/src/sql_connection.dart")));
	sqlResultClass = HandleError(Dart_GetType(dartMssqlLib, Dart_NewStringFromCString("SqlResult"), 0, NULL));
	sqlReturnClass = HandleError(Dart_GetType(dartMssqlLib, Dart_NewStringFromCString("_SqlReturn"), 0, NULL));

	Dart_Handle dartSqlReturn = HandleError(Dart_New(sqlReturnClass, Dart_Null(), 0, NULL));
	sqlConnect(serverName, databaseName, userName, password, authType, &pSession, &errorCount, &errorMessages);
	if (errorCount == 0) {
		registerHandleInGlobalInterfaceTable(pSession, dartSqlReturn, &errorCount, &errorMessages);
	};
	if (errorCount != 0) {
		HandleError(Dart_SetField(dartSqlReturn, Dart_NewStringFromCString("error"), Dart_NewStringFromCString(errorMessages.c_str())));
	}
	Dart_SetReturnValue(arguments, dartSqlReturn);
	Dart_ExitScope();
}

void _executeCommand(Dart_NativeArguments arguments) {
	const char *sqlCommand;
	int64_t handle;
	int	 errorCount = 0;
	std::string errorMessages;
	IDBInitialize* pSession = NULL;

	Dart_EnterScope();
	
	HandleError(Dart_IntegerToInt64(HandleError(Dart_GetNativeArgument(arguments, 0)), &handle));
	HandleError(Dart_StringToCString(HandleError(Dart_GetNativeArgument(arguments, 1)), &sqlCommand));
	Dart_Handle sqlParams = HandleError(Dart_GetNativeArgument(arguments, 2));
	dartMssqlLib = HandleError(Dart_LookupLibrary(Dart_NewStringFromCString("package:dart_mssql/src/sql_connection.dart")));
	sqlResultClass = HandleError(Dart_GetType(dartMssqlLib, Dart_NewStringFromCString("SqlResult"), 0, NULL));
	sqlReturnClass = HandleError(Dart_GetType(dartMssqlLib, Dart_NewStringFromCString("_SqlReturn"), 0, NULL));

	pSession = getHandleFromGlobalInterfaceTable((DWORD)handle, &errorCount, &errorMessages);
	Dart_Handle dartSqlReturn = HandleError(Dart_New(sqlReturnClass, Dart_Null(), 0, NULL));
	Dart_Handle dartSqlResult = HandleError(Dart_New(sqlResultClass, Dart_Null(), 0, NULL));
	if (errorCount == 0) {
		BSTR sqlCmd = UTF8ToBSTR(sqlCommand);
		sqlExecute(pSession, sqlCmd, sqlParams, dartMssqlLib, dartSqlResult, &errorCount, &errorMessages);
	}
	if (errorCount != 0) {
		HandleError(Dart_SetField(dartSqlReturn, Dart_NewStringFromCString("error"), Dart_NewStringFromCString(errorMessages.c_str())));
	}
	else {
		HandleError(Dart_SetField(dartSqlReturn, Dart_NewStringFromCString("result"), dartSqlResult));
	}

	Dart_SetReturnValue(arguments, dartSqlReturn);
	Dart_ExitScope();
}

void _disconnectCommand(Dart_NativeArguments arguments) {
	uint64_t handle;
	int	 errorCount = 0;
	std::string errorMessages;
	IDBInitialize* pSession = NULL;

	Dart_EnterScope();

	HandleError(Dart_IntegerToUint64(HandleError(Dart_GetNativeArgument(arguments, 0)), &handle));
	pSession = getHandleFromGlobalInterfaceTable((DWORD)handle, &errorCount, &errorMessages);
	revokeInterfaceFromGlobalInterfaceTable((DWORD)handle);
	if (pSession) {
		pSession->Release();
	}
	Dart_ExitScope();
}

struct FunctionLookup {
	const char* name;
	Dart_NativeFunction function;
};

FunctionLookup function_list[] = {
	{"_executeCommand", _executeCommand},
	{"_connectCommand", _connectCommand},
	{"_disconnectCommand", _disconnectCommand},
	{NULL, NULL}
};

Dart_NativeFunction nativeResolver(Dart_Handle name, int argc, bool * autoSetupScope) {
	if (!Dart_IsString(name)) {
		return NULL;
	}
	Dart_NativeFunction result = NULL;
	if (autoSetupScope == NULL) {
		return NULL;
	}

	Dart_EnterScope();
	const char* cname;
	HandleError(Dart_StringToCString(name, &cname));

	for (int i = 0; function_list[i].name != NULL; ++i) {
		if (strcmp(function_list[i].name, cname) == 0) {
			*autoSetupScope = true;
			result = function_list[i].function;
			break;
		}
	}

	if (result != NULL) {
		Dart_ExitScope();
		return result;
	}

	Dart_ExitScope();
	return result;
}

DART_EXPORT Dart_Handle dart_mssql_Init(Dart_Handle library) {

	// This should have to be CoUnitialized somewhere. Perhaps when Dart provides a "Done" hook?
	CoInitializeEx(NULL, COINIT_MULTITHREADED | COINIT_DISABLE_OLE1DDE);

	if (Dart_IsError(library)) {
		return library;
	}
	Dart_Handle result = Dart_SetNativeResolver(library, nativeResolver, NULL);
	if (Dart_IsError(result)) {
		return result;
	}

	return Dart_Null();
}

