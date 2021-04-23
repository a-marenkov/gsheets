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

//    /// last row
//    expect(
//      await ws.values.lastRow(length: 2),
//      ['', '2'],
//    );
//
//    expect(
//      await ws.values.lastRow(fromColumn: 3),
//      ['3'],
//    );
//
//    expect(
//      await ws.values.lastRow(length: 1, inRange: true),
//      ['0'],
//    );
//
//    expect(
//      await ws.values.lastRow(length: 1, inRange: false),
//      [''],
//    );
//
//    /// last column
//    expect(
//      await ws.values.lastColumn(length: 2),
//      ['', '3'],
//    );
//
//    expect(
//      await ws.values.lastColumn(fromRow: 2),
//      ['3', '0'],
//    );
//
//    expect(
//      await ws.values.lastColumn(length: 1, inRange: true),
//      ['3'],
//    );
//
//    expect(
//      await ws.values.lastColumn(length: 1, inRange: false),
//      [''],
//    );

//    /// all rows
//    expect(
//      await ws.values.allRows(),
//      [
//        ['1', '2', '3'],
//        ['0', '1', '2', '3'],
//        ['0', '0', '0', '0'],
//        ['', '2', '3'],
//      ],
//    );
//
//    expect(
//      await ws.values.allRows(fill: true),
//      [
//        ['1', '2', '3', ''],
//        ['0', '1', '2', '3'],
//        ['0', '0', '0', '0'],
//        ['', '2', '3', ''],
//      ],
//    );
//
//    expect(
//      await ws.values.allRows(fromRow: 2, fromColumn: 2, length: 2, count: 2),
//      [
//        ['1', '2'],
//        ['0', '0'],
//      ],
//    );
//
//    /// all columns
//    expect(await ws.values.allColumns(), [
//      ['1', '0', '0'],
//      ['2', '1', '0', '2'],
//      ['3', '2', '0', '3'],
//      ['', '3', '0'],
//    ]);
//
//    expect(
//      await ws.values.allColumns(fromRow: 2, fromColumn: 2, length: 2, count: 2),
//      [
//        ['1', '0'],
//        ['2', '0'],
//      ],
//    );

//    /// value
//    expect(await ws.values.value(column: 1, row: 1), '1');
//    expect(await ws.values.value(column: 3, row: 2), '2');
//    expect(await ws.values.value(column: 1, row: 4), '');
//
//    /// valueByKeys
//    expect(await ws.values.valueByKeys(rowKey: 0, columnKey: 2), '1');
//    expect(await ws.values.valueByKeys(rowKey: 0, columnKey: 3), '2');
//    expect(await ws.values.valueByKeys(rowKey: 2, columnKey: 3), null);
//
//    /// insertValue
//    await ws.values.insertValue(2, row: 4, column: 1);
//    expect(await ws.values.value(row: 4, column: 1), '2');
//
//    /// insertValue
//    await ws.values.insertValueByKeys(1, rowKey: 2, columnKey: 2);
//    expect(await ws.values.valueByKeys(rowKey: 2, columnKey: 2), '1');

    /// insertRow
    await ws.values.insertRow(5, [1], fromColumn: 2);
    expect(await ws.values.row(5), ['', '1']);

    /// insertColumn
    await ws.values.insertColumn(1, [2, 3, 4], fromRow: 2);
    expect(await ws.values.column(1), ['1', '2', '3', '4']);

    /// insertColumn
    await ws.values.insertColumn(1, [2, 3, 4, 5], fromRow: 2);
    expect(await ws.values.column(1), ['1', '2', '3', '4', '5']);

    /// rowByKey
    expect(await ws.values.rowByKey(5), ['1']);

    /// columnByKey
    expect(await ws.values.columnByKey(3), ['2', '0', '3']);
  });
}
