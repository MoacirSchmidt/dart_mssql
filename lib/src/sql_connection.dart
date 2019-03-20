import 'dart-ext:dart_mssql';
import 'package:quiver/strings.dart';

String strSlice(String s, [int num = 1]) {
  return isEmpty(s) ? "" : s.substring(0, s.length - 1);
}

String sqlDate(DateTime d) {
  String r = d.toString();
  r = r.substring(0, r.length - 4);
  if (r.endsWith(".")) {
    r += "0";
  }
  return r;
}

class SqlRecord {
  SqlRecord();

  SqlResult _sqlResult;
  List _values;

  @override  //overring noSuchMethod
  // noSuchMethod(Invocation invocation) => 'Got the ${invocation.memberName}';

  /*noSuchMethod(Invocation invocation) { 
    String name = invocation.memberName;
    print('Got the ${name}'); 
    print('Got the ${invocation.memberName}'); 
  }
  */

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

class SqlResult  {
  List columns;
  List rows;
  int rowsAffected = -1;
  int lastIdentity = -1;
  String _error;
  int _rowIndex = 0;
  Map<String, int> _fieldIndexes = Map();

  void _updateFieldIndexes() {
    int i = 0;
    columns?.forEach((c) {
      _fieldIndexes[c] = i;
      i++;
    });
  }

  SqlResult();

  SqlResult.fromJson(Map map) {
    this.columns = List();
    this.rows = List();
    this.rows.add(List());
    map.forEach((k,v) {
      Map<String, dynamic> c = Map();
      c["name"] = k;
      c["type"] = 1; // missing mapping sql types from dart types?
      this.columns.add(c);
      this.rows[0].add(v);
    });
  }

  SqlResult.fromMap(Map map) {
    this.rows = map["rows"];
    this.columns = map["columns"];
    this.rowsAffected = map["rowsAffected"];
    this.lastIdentity = map["lastIdentity"];
    this._fieldIndexes = map["_fieldIndexes"];
    if (this._fieldIndexes == null) {
      _fieldIndexes = Map();
    }
    if (this._fieldIndexes.isEmpty) {
      _updateFieldIndexes();
    }
  }

  int index(String fieldName) {
    return _fieldIndexes[fieldName];
  }

  String values(String fieldName) {
    int fieldIndex = _fieldIndexes[fieldName];
    String ids = "-1,";
    rows.forEach((e) {
      ids += e[fieldIndex].toString() + ",";
    });
    ids = strSlice(ids);
    return ids;
  }
  
  Map asMap() {
    return {
      "columns": columns,
      "rows": rows,
      "rowsAffected": rowsAffected,
      "lastIdentity": lastIdentity,
      "_fieldIndexes": _fieldIndexes,
    };
  }

  String mapToInsert(String tableName, Map<String, dynamic> map,
      [String fieldList]) {
    String fieldNames = "";
    String fieldValues = "";
    if (fieldList != null) {
      fieldList = ",$fieldList,";
    }
    map.forEach((n, v) {
      if (fieldList == null || fieldList.contains(",$n,")) {
        fieldNames += n + ',';
        if (v is String) {
          if (n.indexOf("dat_") == 0) {
            fieldValues +=
                "${v == null ? null : "'" + sqlDate(DateTime.tryParse(v).toUtc()) + "'"},";
          } else {
            fieldValues +=
                "${v == null ? null : "'" + v.replaceAll("'", "''") + "'"},";
          }
        } else {
          fieldValues += "$v,";
        }
      }
    });
    fieldNames = strSlice(fieldNames);
    fieldValues = strSlice(fieldValues);
    return "insert into $tableName($fieldNames) values ($fieldValues)";
  }

  String mapToUpdate(
      String tableName, Map<String, dynamic> map, String whereClause) {
    String fieldValues = "";
    map.forEach((n, v) {
      if (v is String) {
        if (n.indexOf("dat_") == 0) {
          fieldValues +=
              "$n = ${v == null ? null : "'" + sqlDate(DateTime.tryParse(v).toUtc()) + "'"},";
        } else {
          fieldValues +=
              "$n = ${v == null ? null : "'" + v.replaceAll("'", "''") + "'"},";
        }
      } else {
        fieldValues += "$n = $v,";
      }
    });
    fieldValues = strSlice(fieldValues);
    return "update $tableName set $fieldValues where $whereClause";
  }
}

SqlResult _executeCommand(String host, String databaseName, String username, String password, int authType, String sqlCommand, List<dynamic> params, int handle) native '_executeCommand';

int _connectCommand(String host, String databaseName, String username, String password, int authType) native '_connectCommand';

void _disconnectCommand(int handle) native '_disconnectCommand';

class SqlConnection {
  String _username;
  String _password;
  String _host;
  String _databaseName;

  SqlConnection(
      this._host, this._databaseName, this._username, this._password);

  SqlResult execute(String sqlCommand, [List<dynamic> params]) {
    int handle = _connectCommand(_host, _databaseName, _username, _password, 0);
    print(handle);
    SqlResult result = _executeCommand(_host, _databaseName, _username, _password, 0, sqlCommand, params, handle);
    _disconnectCommand(handle);
    if (isNotEmpty(result._error))
      throw result._error;
    result._updateFieldIndexes();
    return result;
  }
}
