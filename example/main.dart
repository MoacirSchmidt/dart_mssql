import 'dart:io';
import 'package:dart_mssql/dart_mssql.dart';

void main() async {
  connection.open();
  String cmd = "select * from  nacionalidade";
  
  SqlResult result = connection.execute(cmd);
  print(connection.lastIdentity());
  result.rows.forEach((e) {
    print("${e.id_nacionalidade}: ${e.nom_nacionalidade}");
  });
  print("end of printing.");
  connection.close();
  stdin.readLineSync();
}
