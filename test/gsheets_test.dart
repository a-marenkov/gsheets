// your google auth credentials
import 'package:gsheets/gsheets.dart';
import 'package:test/test.dart';

import 'env.dart';

void main() async {
  final gsheets = GSheets(TestsConfig.credentials);
  final ssheet = await gsheets.spreadsheet(TestsConfig.spreadsheetId);

  if (ssheet.worksheetByTitle('tests') != null) {
    await ssheet.deleteWorksheet(ssheet.worksheetByTitle('tests')!);
  }

  final ws = await ssheet.addWorksheet('tests');

  test('insert and read', () async {
    await ws.values.insertRow(1, [1, 2, 3]);

    expect(
      await ws.values.row(1),
      ['1', '2', '3'],
    );

    await ws.values.insertRow(2, [1, 2, 3], fromColumn: 2);

    expect(
      await ws.values.row(2),
      ['', '1', '2', '3'],
    );

    expect(
      await ws.values.column(2),
      ['2', '1'],
    );

    expect(
      await ws.values.column(2, fromRow: 2),
      ['1'],
    );

    await ws.values.insertColumn(1, [0], fromRow: 2);

    expect(
      await ws.values.column(1),
      ['1', '0'],
    );

    expect(
      await ws.values.rowByKey(1),
      ['2', '3', ''],
    );

    expect(
      await ws.values.rowByKey(1, fromColumn: 3),
      ['3'],
    );

    expect(
      await ws.values.rowByKey(1, length: 1),
      ['2'],
    );

  });
}
