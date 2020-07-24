import 'dart:convert';
import 'dart:math';

import 'package:http/http.dart';

import 'gsheets.dart';

const DIMEN_ROWS = 'ROWS';
const DIMEN_COLUMNS = 'COLUMNS';
const defaultRowsCount = 1000;
const defaultColumnCount = 26;
const gsheetsCellsLimit = 5000000;

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

String parseKey(dynamic key, [String type = '']) {
  final k = key is String ? key : key?.toString();
  if (isNullOrEmpty(k)) throw GSheetsException('invalid $type key ($key)');
  return k;
}

String parseMapToKey(dynamic key) {
  return key is String ? key : key?.toString();
}

String parseValue(dynamic value) =>
    value is String ? value : value?.toString() ?? '';

void checkValues(dynamic values) {
  if (isNullOrEmpty(values)) throw GSheetsException('invalid values ($values)');
}

void checkNotNested(List<dynamic> values) {
  if (values is List<List>) {
    throw GSheetsException('invalid values type (${values.runtimeType})');
  }
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

int inRangeIndex(List<List<String>> lists, int offset) {
  int index;
  for (var i = lists.length - 1; i > 0; i--) {
    if (lists[i].length > offset - 1) {
      break;
    } else {
      index = i;
    }
  }
  return index ?? lists.length;
}

int maxLength(List<List> lists, [int atLeast = 0]) {
  var length = atLeast;
  for (var list in lists) {
    if (list.length > length) length = list.length;
  }
  return length;
}

bool isNullOrEmpty(dynamic data) => data == null || data.isEmpty;

bool gridSheetsFilter(json) => json['properties']['sheetType'] == 'GRID';
