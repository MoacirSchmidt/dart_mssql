import 'dart-ext:dart_mssql';
import 'package:meta/meta.dart';

import 'package:string_utils/string_utils.dart';
import 'package:sql_result/sql_result.dart';
import 'package:sql_utils/sql_utils.dart';

String _strSlice(String s, [int num = 1]) {
  return isEmpty(s) ? "" : s.substring(0, s.length - 1);
}

class _SqlReturn {
  int handle;
  String error;
  SqlResult result;
}

_SqlReturn _connectCommand(String host, String databaseName, String username, String passwordword, int authType) native '_connectCommand';

_SqlReturn _executeCommand(int handle, String sqlCommand, List<dynamic> params) native '_executeCommand';

_SqlReturn _disconnectCommand(int handle) native '_disconnectCommand';

/// This is the primary type of this library, a connection is responsible for connecting to databases and executing queries. 
class SqlConnection {

  /// Name of database server
  String host;

  /// Name of database 
  String db;

  /// user name used to connect
  String user;

  /// password used to connect 
  String password;

  int _handle = 0;

  SqlConnection({
      @required this.host, @required this.db, this.user, this.password}) {
    _SqlReturn r = _connectCommand(host, db, user, password, 0);
    if (isNotEmpty(r.error)) {
      throw r.error;
    }
    _handle = r.handle;
  }

  void _checkHandle() {
    if (_handle <= 0) {
      throw "Not connected";
    }
  }

  /// Executes a sql command with optional params 
  SqlResult execute(String sqlCommand, [List<dynamic> params]) {
    _checkHandle();
    _SqlReturn r = _executeCommand(_handle, sqlCommand, params);
    if (isNotEmpty(r.error)) {
      throw r.error;
    }
    r.result.updateFieldIndexes();
    return r.result;
  }

  /// Executes a sql command that returns one single row (allow optional params)
  SqlRow selectOne(String sqlCommand, [List<dynamic> params]) {
    _checkHandle();
    _SqlReturn r = _executeCommand(_handle, sqlCommand, params);
    if (isNotEmpty(r.error)) {
      throw r.error;
    }
    r.result.updateFieldIndexes();
    if (r.result.rows.isEmpty) {
      return null;
    } else {
      return r.result.rows.first;
    }
  }

  /// Returns the identity column value of last inserted row
  int lastIdentity() {
    SqlResult r = execute("select @@identity as idty");
    if (r.rows != null && r.rows.isNotEmpty) {
      return r.rows.first.idty.round();
    } else {
      return null;
    }
  }

  /// Inserts [row] into [tableName]. 
  /// Returns the number of rows inserted.
  /// if [onlyColumns] is specified, only this columns will be used from [row] map.
  /// if [excludedColumns] is specified, all columns from [row] map will be used EXCEPT these ones.
  int insert(String tableName, Map<String, dynamic> row, {List<String> onlyColumns, List<String> excludedColumns}) {
    String fieldNames = "";
    String fieldValues = "";
    row.forEach((n, v) {
      if ((onlyColumns == null || onlyColumns.contains(n)) && (excludedColumns == null || !excludedColumns.contains(n))) {
        fieldNames += n + ',';
        if (v is DateTime) {
          fieldValues +=
              "${v == null ? null : "'" + sqlDateTime(v) + "'"},";
        } else if (v is String) {
          fieldValues +=
              "${v == null ? null : "'" + v.replaceAll("'", "''") + "'"},";
        } else {
          fieldValues += "$v,";
        }
      };
    });
    fieldNames = _strSlice(fieldNames);
    fieldValues = _strSlice(fieldValues);
    SqlResult r = execute("insert into $tableName($fieldNames) values ($fieldValues)");
    return r.rowsAffected;
  }
  
  /// Updates [row] into [tableName]. 
  /// Returns the number of rows updated
  /// [where] and [whereArgs] will be used as update where clause
  /// if [onlyColumns] is specified, only this columns will be used from [row] map
  /// if [excludedColumns] is specified, all columns from [row] map will be used EXCEPT these ones
  update(String tableName, Map<String, dynamic> row, String where, List<dynamic> whereArgs, {List<String> onlyColumns, List<String> excludedColumns}) {
    assert(isNotEmpty(where) && whereArgs != null && whereArgs.isNotEmpty);
    String fieldValues = "";
    row.forEach((n, v) {
      if ((onlyColumns == null || onlyColumns.contains(n)) && (excludedColumns == null || !excludedColumns.contains(n))) {
        fieldValues += n + '=';
        if (v is DateTime) {
          fieldValues +=
              "${v == null ? null : "'" + sqlDateTime(v) + "'"},";
        } else if (v is String) {
          fieldValues +=
              "${v == null ? null : "'" + v.replaceAll("'", "''") + "'"},";
        } else {
          fieldValues += "$v,";
        }
      }
    });
    fieldValues = _strSlice(fieldValues);
    SqlResult r = execute("update $tableName set $fieldValues where $where", whereArgs);
    return r.rowsAffected;
    }

  /// Delete [row] from [tableName]. 
  /// Returns the number of rows deleted
  /// [where] and [whereArgs] will be used as delete where clause
  delete(String tableName, String where, List<dynamic> whereArgs) {
    assert(isNotEmpty(where) && whereArgs != null && whereArgs.isNotEmpty);
    SqlResult r = execute("delete $tableName where $where", whereArgs);
    return r.rowsAffected;
  }

  /// Closes the connection. Remember to allways close connections afer use
  void close() {
    if (_handle > 0) {
      _disconnectCommand(_handle);
    }
    _handle = 0;
  }
}

