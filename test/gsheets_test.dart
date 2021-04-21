// your google auth credentials
import 'package:gsheets/gsheets.dart';
import 'package:test/test.dart';

import 'env.dart';

void main() async {
  final gsheets = GSheets(TestsConfig.credentials);
  final ssheet = await gsheets.spreadsheet(TestsConfig.spreadsheetId);

//  if (ssheet.worksheetByTitle('tests') != null) {
//    await ssheet.deleteWorksheet(ssheet.worksheetByTitle('tests')!);
//  }
//
//  final ws = await ssheet.addWorksheet('tests');

  final ws = ssheet.worksheetByTitle('tests')!;

  test('insert and read', () async {
//    await ws.values.insertRow(1, [1, 2, 3]);
//
//    expect(
//      await ws.values.row(1),
//      ['1', '2', '3'],
//    );
//
//    await ws.values.insertRow(2, [1, 2, 3], fromColumn: 2);
//
//    expect(
//      await ws.values.row(2),
//      ['', '1', '2', '3'],
//    );
//
//    expect(
//      await ws.values.column(2),
//      ['2', '1'],
//    );
//
//    expect(
//      await ws.values.column(2, fromRow: 2),
//      ['1'],
//    );
//
//    await ws.values.insertColumn(1, [0], fromRow: 2);
//
//    expect(
//      await ws.values.column(1),
//      ['1', '0'],
//    );
//
//    expect(
//      await ws.values.rowByKey(1),
//      ['2', '3'],
//    );
//
//    expect(
//      await ws.values.rowByKey(1, fromColumn: 3),
//      ['3'],
//    );
//
//    expect(
//      await ws.values.rowByKey(1, length: 1),
//      ['2'],
//    );
//
//    await ws.values.insertRow(3, [0, 0, 0, 0]);
//
//    expect(
//      await ws.values.columnByKey(1),
//      ['0', '0'],
//    );
//
//    expect(
//      await ws.values.columnByKey(2, fromRow: 3),
//      ['0'],
//    );
//
//    expect(
//      await ws.values.columnByKey(2, length: 1),
//      ['1'],
//    );
//
//    expect(
//      await ws.values.lastRow(),
//      ['0', '0', '0', '0'],
//    );
//
//    await ws.values.appendRow([2,3], fromColumn: 2);

    /// last row
    expect(
      await ws.values.lastRow(length: 2),
      ['', '2'],
    );

    expect(
      await ws.values.lastRow(fromColumn: 3),
      ['3'],
    );

    expect(
      await ws.values.lastRow(length: 1, inRange: true),
      ['0'],
    );

    expect(
      await ws.values.lastRow(length: 1, inRange: false),
      [''],
    );

    /// last column
    expect(
      await ws.values.lastColumn(length: 2),
      ['', '3'],
    );

    expect(
      await ws.values.lastColumn(fromRow: 2),
      ['3', '0'],
    );

    expect(
      await ws.values.lastColumn(length: 1, inRange: true),
      ['3'],
    );

    expect(
      await ws.values.lastColumn(length: 1, inRange: false),
      [''],
    );
  });
}
