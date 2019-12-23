import 'dart:convert';
import 'dart:math';

import 'package:http/http.dart';

import 'gsheets.dart';

const DIMEN_ROWS = 'ROWS';
const DIMEN_COLUMNS = 'COLUMNS';

final int _char_a = 'A'.codeUnitAt(0);

String getColumnLetter(int index) {
  if (index < 1) throw GSheetsException('invalid index ($index)');
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

void checkCR(int column, int row) {
  checkC(column);
  checkR(row);
}

void checkC(int column) {
  if ((column ?? -1) < 1) throw GSheetsException('invalid column ($column)');
}

void checkR(int row) {
  if ((row ?? -1) < 1) throw GSheetsException('invalid row ($row)');
}

void checkI(int index) {
  if ((index ?? -1) < 1) throw GSheetsException('invalid index ($index)');
}

void checkL(int length) {
  if ((length ?? -1) < 1) throw GSheetsException('invalid length ($length)');
}

void checkV(dynamic values) {
  if (isNullOrEmpty(values)) throw GSheetsException('invalid values ($values)');
}

void checkM(dynamic first, dynamic second) {
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

void mapKeysToValues<V>(
  List<String> keys,
  List<V> values,
  Map<String, V> map,
  V defaultTo,
  V Function(int index) wrap,
) {
  var index = 0;
  var length = values.length;
  if (wrap == null) {
    for (var key in keys) {
      map[key] = index < length ? values[index] : defaultTo;
      index++;
    }
  } else {
    for (var key in keys) {
      map[key] = index < length ? values[index] : wrap(index);
      index++;
    }
  }
}

int maxLength(List<List<Object>> data) {
  var len = 0;
  data.forEach((list) {
    len = max(len, list.length);
  });
  return len;
}

class Tuple<A, B> {
  final A first;
  final B second;

  const Tuple(this.first, this.second);

  @override
  String toString() => 'first($first)|second($second)';
}

String getOrEmpty(List<String> list, [int index = 0]) =>
    list == null || list.length < index + 1 ? '' : list[index];

bool isNullOrEmpty(dynamic data) => data == null || data.isEmpty;
