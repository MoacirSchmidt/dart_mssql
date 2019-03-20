#pragma once

//ROUNDUP on all platforms pointers must be aligned properly
constexpr auto ROUNDUP_AMOUNT = 8;
#define ROUNDUP_(size,amount)	(((DBBYTEOFFSET)(size)+((amount)-1))&~((DBBYTEOFFSET)(amount)-1))
#define ROUNDUP(size)			ROUNDUP_(size, ROUNDUP_AMOUNT)

Dart_Handle HandleError(Dart_Handle handle);

BSTR UTF8ToBSTR(const char* utf8Value);

char* WideCharToUTF8(const wchar_t* value);

BSTR intToBSTR(int number);

void addError(int* errorCount, std::string* errorMessages, const std::string errorMessage);

HRESULT oleCheck(HRESULT hr, int* errorCount, std::string* errorMessages);

HRESULT memoryCheck(HRESULT hr, void *pv, int* errorCount, std::string* errorMessages);

HRESULT createShortDataAccessor(IUnknown* pUnkRowset, HACCESSOR* phAccessor, DBORDINAL* pcBindings, DBBINDING** prgBindings, DBORDINAL cColumns, DBCOLUMNINFO *rgColumnInfo, DBORDINAL* pcbRowSize, int* errorCount, std::string* errorMessages);

HRESULT createLargeDataAcessors(IUnknown* pUnkRowset, HACCESSOR** phAccessors, DBORDINAL* pcBindings, DBBINDING** prgBindings, DBORDINAL cColumns, DBCOLUMNINFO *rgColumnInfo, int* errorCount, std::string* errorMessages);

HRESULT createParamsAccessor(IUnknown* pUnkRowset, HACCESSOR* phAccessor, DBCOUNTITEM sqlParamsCount, Dart_Handle sqlParams, DBORDINAL* pcbRowSize, void** ppParamData, int* errorCount, std::string* errorMessages);

void freeBindings(DBORDINAL	cBindings, DBBINDING *rgBindings);

HRESULT sqlConnect(const char* serverName, const char* dbName, const char* userName, const char* password, int64_t authType, IDBInitialize** ppInitialize, int* errorCount, std::string* errorMessages);

void sqlExecute(IDBInitialize* pInitialize, const LPCOLESTR sqlCommand, Dart_Handle sqlParams, Dart_Handle dartMssqlLib, Dart_Handle dartSqlResult, int* errorCount, std::string* errorMessages);

