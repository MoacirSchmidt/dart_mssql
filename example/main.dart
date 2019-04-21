import 'dart:io';

import 'package:dart_mssql/dart_mssql.dart';

// clas "Client" for ORM example
class Client {
  int client_id;
  String client_name;
  List<Invoice> invoices;

  Client.fromJson(Map<String, dynamic> json) {
    client_id = json['client_id'];
    client_name = json['client_name'];
  }  

  Map<String, dynamic> toJson() {  
    return {
      'client_id': client_id,
      'client_name': client_name,
    };
  }
}

// clas "Invoice" for ORM example
class Invoice {
  int client_id;
  int inv_number;

  Invoice.fromJson(Map<String, dynamic> json) {
    client_id = json['client_id'];
    inv_number = json['inv_number'];
  }  

  Map<String, dynamic> toJson() {  
    return {
      'client_id': client_id,
      'inv_number': inv_number,
    };
  }
}

void main() {
  // Establish a connection
  SqlConnection connection = SqlConnection(host:"SERVERNAME", db:"DBNAME", user:"USERNAME", password:"PASSWORD");

  // Querying several rows
  String cmd = "select id_nacionalidade,nom_nacionalidade from nacionalidade where id_nacionalidade>?"; // parameters binding!  
  SqlResult result = connection.execute(cmd,[4]);
  result.rows.forEach((e) {
    print("${e.id_nacionalidade}");
  });

  // Querying one row
  cmd = "select id_nacionalidade,nom_nacionalidade from nacionalidade where id_nacionalidade=?"; 
  dynamic row = connection.selectOne(cmd,[4]); // "dynamic" var is important...
  print(row.id_nacionalidade); // ...to allow accessing fields by name

  // Bonus: How to make a master/detail ORM query:
  SqlResult master = connection.execute("select client_id, client_name from client");
  SqlResult detail = connection.execute("select client_id, inv_number from invoice");
  master.rows.forEach((r) {
    Client client = Client.fromJson(r.toJson());
    client.invoices = detail.rows.where((e) => e.client_id == client.client_id).map((e) => Invoice.fromJson(e.toJson())).toList();
  });  
  
  print("end of printing.");
  connection.close();
  stdin.readLineSync();
}
