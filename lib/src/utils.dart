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
  if (from == 0 && length == -1) return list;
  final to = length < 1 ? list.length : min(list.length, from + length);
  return list.sublist(min(list.length, from), to);
}

T get<T>(List<T> list, {int at = 0, T or}) =>
    (list?.length ?? 0) > at ? list[at] : or;

int whereFirst(List<List<String>> list, String key) =>
    list.indexWhere((it) => get(it) == key);

String getOrEmpty(List<String> list, [int at = 0]) => get(list, at: at, or: '');

bool isNullOrEmpty(dynamic data) => data == null || data.isEmpty;
