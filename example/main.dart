import 'dart:io';
import 'package:dart_mssql/dart_mssql.dart';

void main() async {
  SqlConnection connection = SqlConnection("SERVERNAME", "DATABASENAME", "USERNAME", "PASSWORD");
  String cmd = "select * from nacionalidade where nom_nacionalidade like ?";
  
  SqlResult result = connection.execute(cmd,['%Bras%']);
  result.rows.forEach((e) {
    print(e.id_nacionalidade);
    print(e.nom_nacionalidade);
  });
  print("end of printing.");
  stdin.readLineSync();
}
