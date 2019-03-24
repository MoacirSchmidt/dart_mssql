import 'dart:io';
import 'package:dart_mssql/dart_mssql.dart';

void main() async {
  SqlConnection connection = SqlConnection(host:"SERVERNAME", db:"DBNAME", user:"USERNAME", password:"PASSWORD");
  String cmd = "select email from usuario where id_usuario=?";
  
  SqlResult result = connection.execute(cmd,[4]);
  result.rows.forEach((e) {
    print("${e.email}");
  });
  print("end of printing.");
  connection.close();
  stdin.readLineSync();
}
