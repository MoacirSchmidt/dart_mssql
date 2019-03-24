import 'dart-ext:dart_mssql';
import 'package:meta/meta.dart';
import 'package:quiver/strings.dart';

import 'package:dart_mssql/recase.dart';

/// Convert a DateTime to a string suitable by MS-SQL commands 
String sqlDateTime(DateTime d) {
  String r = d.toString();
  r = r.substring(0, r.length - 4);
  if (r.endsWith(".")) {
    r += "0";
  }
  return r;
}

String _strSlice(String s, [int num = 1]) {
  return isEmpty(s) ? "" : s.substring(0, s.length - 1);
}

/// Each row returned from a query
class SqlRow {

  SqlRow();

  SqlResult _sqlResult;
  List _values;

  dynamic operator [](int index) => _values[index];

  Map<String, dynamic> toJson({recaseKey:recaseKeyNone}) {
    Map<String, dynamic> map = Map();
    int i = 0;
    _sqlResult?._fieldIndexes?.forEach((k, v) {
      map[recaseKey(k)] = _values[i++];
    });
    return map;
  }

  @override  
  noSuchMethod(Invocation invocation) {
    if (invocation.isGetter) {
      String name = invocation.memberName.toString().substring(8); // Symbol("columnName")
      name = name.substring(0,name.length-2);
      var i = _sqlResult._fieldIndexes[name];
      if (i != null) {
        return _values[i];
      }
    }
    return super.noSuchMethod(invocation);
  }
}

/// Result set returned from a query
class SqlResult  {

  /// List<String> with column names
  List<String> columns;

  /// List of [SqlRow] objects with all returned rows from a query
  List rows;

  /// Number of rows affected by last insert, update or delete command
  int rowsAffected = -1;
  Map<String, int> _fieldIndexes = Map();

  void _updateFieldIndexes() {
    int i = 0;
    columns?.forEach((c) {
      _fieldIndexes[c] = i;
      i++;
    });
  }

  SqlResult();
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
    if (isNotEmpty(r.error))
      throw r.error;
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
    if (isNotEmpty(r.error))
      throw r.error;
    r.result._updateFieldIndexes();
    return r.result;
  }

  /// Returns the identity column value of last inserted row
  int lastIdentity() {
    SqlResult r = execute("select @@identity");
    if (r.rows != null && r.rows.isNotEmpty) {
      return r.rows[0]._values[0];
    } else
      return null;
  }

  /// Inserts row into tableName. 
  /// Returns the number of rows inserted
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

  /// Closes the connection. Remember to allways close connections afer use
  void close() {
    if (_handle > 0) {
      _disconnectCommand(_handle);
    }
    _handle = 0;
  }
}
