import 'dart:convert';
import 'dart:math';

import 'package:http/http.dart';

import 'gsheets.dart';

const dimenRows = 'ROWS';
const dimenColumns = 'COLUMNS';
const defaultRowsCount = 1000;
const defaultColumnCount = 26;
const gsheetsCellsLimit = 5000000;

void checkIndex(String name, int value) =>
    except(value < 1, 'invalid $name ($value)');

String parseKey(Object str, [String type = '']) {
  final key = str is String ? str : str.toString();
  except(key.isEmpty, 'invalid $type key ($str)');
  return key;
}

String parseString(Object? str, [String defaultValue = '']) {
  if (str == null) return defaultValue;
  return str is String ? str : str.toString();
}

String? parseStringOrNull(Object? str) {
  if (str == null) return null;
  return parseKey(str);
}

void checkValues(List<Object?> values) =>
    except(values.isEmpty, 'invalid values ($values)');

void checkNotNested(List values) =>
    except(values is List<List>, 'invalid values type (${values.runtimeType})');

void checkMap(Map map) => except(map.isEmpty, 'invalid map ($map)');

void checkMaps(List<Map> maps) => except(maps.isEmpty, 'invalid maps ($maps)');

void checkMapTo(first, second) =>
    except(first == second, 'cannot map $first to $second');

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
  for (final key in keys) {
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

T? get<T>(List<T> list, {int at = 0, T? or}) =>
    list.length > at ? list[at] : or;

String getOrEmpty(List<String> list, [int at = 0]) =>
    get(list, at: at, or: '')!;

int whereFirst(List<List<String>> lists, String key) =>
    lists.indexWhere((list) => get<String>(list) == key);

int inRangeIndex(List<List<String>> lists, int offset) {
  int? index;
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
  for (final list in lists) {
    if (list.length > length) length = list.length;
  }
  return length;
}

void appendIfShorter<T>(
  List<List<T>> lists,
  int length,
  T appendix,
) {
  for (final list in lists) {
    final dif = length - list.length;
    if (dif > 0) {
      list.addAll(List.generate(dif, (_) => appendix));
    }
  }
}

bool gridSheetsFilter(json) => json['properties']['sheetType'] == 'GRID';

extension StringX on String {
  Uri toUri() => Uri.parse(this);
}

extension StringNX on String? {
  bool get isNullOrEmpty => this?.isEmpty ?? true;
}
