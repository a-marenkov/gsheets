import 'dart:convert';
import 'dart:math';

import 'package:http/http.dart';

import 'gsheets.dart';

const DIMEN_ROWS = 'ROWS';
const DIMEN_COLUMNS = 'COLUMNS';

final int _char_a = 'A'.codeUnitAt(0);

String getColumnLetter(int index) {
  check('index', index);
  var number = index - 1;
  var remainder = number % 26;
  var label = String.fromCharCode(_char_a + remainder);
  number = number ~/ 26;
  while (number > 0) {
    var remainder = number % 26 - 1;
    label = '${String.fromCharCode(_char_a + remainder)}$label';
    number = number ~/ 26;
  }
  return label;
}

void check(String name, int value) {
  if ((value ?? -1) < 1) throw GSheetsException('invalid $name ($value)');
}

void checkKey(String key) {
  if (isNullOrEmpty(key)) throw GSheetsException('invalid key ($key)');
}

void checkValues(dynamic values) {
  if (isNullOrEmpty(values)) throw GSheetsException('invalid values ($values)');
}

void checkMap(dynamic values) {
  if (isNullOrEmpty(values)) throw GSheetsException('invalid map ($values)');
}

void checkMapTo(dynamic first, dynamic second) {
  if (first == second) throw GSheetsException('cannot map $first to $second');
}

void except(bool check, String cause) {
  if (check) throw GSheetsException(cause);
}

void checkResponse(Response response) {
  if (response.statusCode != 200) {
    final msg = (jsonDecode(response.body)['error'] ?? const {})['message'];
    throw GSheetsException(msg ?? response.body);
  }
}

Map<String, V> mapKeysToValues<V>(
  List<String> keys,
  List<String> values,
  V Function(int index, dynamic value) wrap,
) {
  final map = <String, V>{};
  var index = 0;
  var length = values.length;
  for (var key in keys) {
    map[key] = index < length ? wrap(index, values[index]) : wrap(index, null);
    index++;
  }
  return map;
}

List<String> extractSublist(
  List<String> list, {
  int from = 0,
  int length = -1,
}) {
  if (from == 0 && length < 1) return list;
  final start = min(from, list.length);
  final end = length < 1 ? list.length : min(from + length, list.length);
  return list.sublist(start, end);
}

T get<T>(List<T> list, {int at = 0, T or}) {
  return (list?.length ?? 0) > at ? list[at] : or;
}

String getOrEmpty(List<String> list, [int at = 0]) {
  return get(list, at: at, or: '');
}

int whereFirst(List<List<String>> lists, String key) {
  return lists.indexWhere((list) => get<String>(list) == key);
}

int maxLength(List<List> lists, [int atLeast = 0]) {
  var length = atLeast;
  for (var list in lists) {
    if (list.length > length) length = list.length;
  }
  return length;
}

bool isNullOrEmpty(dynamic data) => data == null || data.isEmpty;
