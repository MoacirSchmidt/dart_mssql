import 'package:recase/recase.dart';

String recaseKeyNone(String k) => k;
  
String recaseKeyCamelCase(String k) => ReCase(k).camelCase;

String recaseKeySnakeCase(String k) => ReCase(k).snakeCase; 

Map<String, dynamic> recaseMap(Map<String, dynamic> map, [recaseKey=recaseKeyNone]) {
  if (recaseKey == recaseKeyNone) {
    return map;
  }
  Map<String, dynamic> result = Map();
  map.forEach((k, v) {
    result[recaseKey(k)] = v;
  });
  return result;
}
