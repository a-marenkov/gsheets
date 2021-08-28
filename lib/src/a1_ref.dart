/// Utility class for resolving a1 notation
class A1Ref {
  static final int _charCodeUnitA = 'A'.codeUnitAt(0);
  static final _regExp = RegExp(
    r'^([A-Z]+)([0-9]+)$',
    caseSensitive: false,
  );

  /// Cell reference in A1 notation
  final String value;
  late final RegExpMatch _match = _regExp.firstMatch(value)!;

  /// Cell column index
  late final int column = getColumnIndex(_match.group(1)!);

  /// Cell row index
  late final int row = int.parse(_match.group(2)!);

  /// Creates [A1Ref]
  /// Can be used to resolve indices of cell reference in A1 notation
  A1Ref(String value)
      : assert(
          _regExp.hasMatch(value.trim()),
          'invalid A1 notation reference ($value)',
        ),
        value = value.toUpperCase().trim();

  /// Converts [index] into A1 notation letter label
  /// 1 -> A
  /// 2 -> B
  /// 27 -> AA
  static String getColumnLabel(int index) {
    final res = <String>[];
    var block = index - 1;
    while (block >= 0) {
      res.insert(0, String.fromCharCode((block % 26) + _charCodeUnitA));
      block = block ~/ 26 - 1;
    }
    return res.join();
  }

  /// Converts A1 notation letter [label] into column index
  /// A -> 1
  /// B -> 2
  /// AA -> 27
  static int getColumnIndex(String label) {
    final chars = label.split('');
    var res = 0;
    for (final char in chars) {
      res *= 26;
      res += 1 + char.codeUnitAt(0) - _charCodeUnitA;
    }
    return res;
  }
}
