// dart_mssql.cpp : Defines the exported functions for the DLL application.
//
#include <comdef.h>     
#include <comutil.h>
#include <map>
#include <list>
#include <exception>
#include "msoledbsql.h"
#include "dart_api.h"
#include "dart_native_api.h"
#include "dart_mssql.h"

Dart_Handle HandleError(Dart_Handle handle) {
	if (Dart_IsError(handle)) Dart_PropagateError(handle);
	return handle;
}

Dart_Handle	dartMssqlLib = HandleError(Dart_LookupLibrary(Dart_NewStringFromCString("package:dart_mssql/src/sql_connection.dart")));
Dart_Handle	sqlResultClass = HandleError(Dart_GetType(dartMssqlLib, Dart_NewStringFromCString("SqlResult"), 0, NULL));

class classExecute {
private:
	int	 errorCount = 0;
	std::string errorMessages;
public:
	classExecute(const char* serverName, const char* dbName, const char* userName, const char* password, int64_t authType, const char* sqlCommand, Dart_Handle sqlParams, Dart_Handle* result) {
		CoInitializeEx(NULL, COINIT_MULTITHREADED);
		IDBInitialize* pSession = NULL;
		Dart_Handle	dartSqlResult = HandleError(Dart_New(sqlResultClass, Dart_Null(), 0, NULL));
		*result = dartSqlResult;
		sqlConnect(serverName, dbName, userName, password, authType, &pSession, &errorCount, &errorMessages);
		if (errorCount == 0) {
			BSTR sqlCmd = UTF8ToBSTR(sqlCommand);
			sqlExecute(pSession, sqlCmd, sqlParams, dartMssqlLib, dartSqlResult, &errorCount, &errorMessages);
			SysFreeString(sqlCmd);
			//if (jsonResult["rowsAffected"] > 0) { // melhorar isto pois está sendo executado em updates e deletes?
			//	char s1[20] = "select @@identity";
			//	BSTR s2 = UTF8ToBSTR(s1);
				//json lastIdentity = sqlExecute(pSession, s2, NULL, &errorCount, &errorMessages);
				//jsonResult["lastIdentity"] = lastIdentity["rows"][0][0];
			//	SysFreeString(s2);
			//}
		}
		if (errorCount != 0) {
			HandleError(Dart_SetField(dartSqlResult, Dart_NewStringFromCString("_error"), Dart_NewStringFromCString(errorMessages.c_str())));
		}
	}
	~classExecute() {
		CoUninitialize();
	}
};

void _connectCommand(Dart_NativeArguments arguments) {
	const char *serverName;
	const char *databaseName;
	const char *userName;
	const char *password;
	int64_t authType;
	Dart_Handle result;
	int	 errorCount = 0;
	std::string errorMessages;
	IDBInitialize* pSession = NULL;

	Dart_EnterScope();
	HandleError(Dart_StringToCString(HandleError(Dart_GetNativeArgument(arguments, 0)), &serverName));
	HandleError(Dart_StringToCString(HandleError(Dart_GetNativeArgument(arguments, 1)), &databaseName));
	HandleError(Dart_StringToCString(HandleError(Dart_GetNativeArgument(arguments, 2)), &userName));
	HandleError(Dart_StringToCString(HandleError(Dart_GetNativeArgument(arguments, 3)), &password));
	HandleError(Dart_IntegerToInt64(HandleError(Dart_GetNativeArgument(arguments, 4)), &authType));

	CoInitializeEx(NULL, COINIT_MULTITHREADED);
	sqlConnect(serverName, databaseName, userName, password, authType, &pSession, &errorCount, &errorMessages);
	if (errorCount == 0) {
		result = HandleError(Dart_NewIntegerFromUint64((uintptr_t)pSession));
	} else {
		throw errorMessages;
	}

	Dart_SetReturnValue(arguments, result);
	Dart_ExitScope();
}

void _executeCommand(Dart_NativeArguments arguments) {
	const char *serverName;
	const char *databaseName;
	const char *userName;
	const char *password;
	const char *sqlCommand;
	int64_t authType;
	int64_t handle;
	int	 errorCount = 0;
	std::string errorMessages;
	IDBInitialize* pSession = NULL;

	Dart_EnterScope();
	HandleError(Dart_StringToCString(HandleError(Dart_GetNativeArgument(arguments, 0)), &serverName));
	HandleError(Dart_StringToCString(HandleError(Dart_GetNativeArgument(arguments, 1)), &databaseName));
	HandleError(Dart_StringToCString(HandleError(Dart_GetNativeArgument(arguments, 2)), &userName));
	HandleError(Dart_StringToCString(HandleError(Dart_GetNativeArgument(arguments, 3)), &password));
	HandleError(Dart_IntegerToInt64(HandleError(Dart_GetNativeArgument(arguments, 4)), &authType));
	HandleError(Dart_StringToCString(HandleError(Dart_GetNativeArgument(arguments, 5)), &sqlCommand));
	Dart_Handle sqlParams = HandleError(Dart_GetNativeArgument(arguments, 6));
	HandleError(Dart_IntegerToInt64(HandleError(Dart_GetNativeArgument(arguments, 7)), &handle));

	//classExecute execute(serverName, databaseName, userName, password, authType, sqlCommand, sqlParams, &result);
	Dart_Handle	dartMssqlLib = HandleError(Dart_LookupLibrary(Dart_NewStringFromCString("package:dart_mssql/src/sql_connection.dart")));
	Dart_Handle	sqlResultClass = HandleError(Dart_GetType(dartMssqlLib, Dart_NewStringFromCString("SqlResult"), 0, NULL));

	pSession = reinterpret_cast<IDBInitialize*>(handle);
	Dart_Handle	dartSqlResult = HandleError(Dart_New(sqlResultClass, Dart_Null(), 0, NULL));
	BSTR sqlCmd = UTF8ToBSTR(sqlCommand);
	sqlExecute(pSession, sqlCmd, sqlParams, dartMssqlLib, dartSqlResult, &errorCount, &errorMessages);


	Dart_SetReturnValue(arguments, dartSqlResult);
	Dart_ExitScope();
}

void _disconnectCommand(Dart_NativeArguments arguments) {
	int64_t handle;
	IDBInitialize* pSession = NULL;

	Dart_EnterScope();
	HandleError(Dart_IntegerToInt64(HandleError(Dart_GetNativeArgument(arguments, 0)), &handle));
	pSession = reinterpret_cast<IDBInitialize*>(handle);
	if (pSession) {
		pSession->Release();
	}
	CoUninitialize();
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
	if (Dart_IsError(library)) {
		return library;
	}
	Dart_Handle result = Dart_SetNativeResolver(library, nativeResolver, NULL);
	if (Dart_IsError(result)) {
		return result;
	}
	return Dart_Null();
}

