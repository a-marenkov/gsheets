import 'dart:convert';
import 'package:googleapis/sheets/v4.dart' as v4;
import 'package:googleapis_auth/auth_io.dart';
import 'package:http/http.dart';
import 'package:meta/meta.dart';
import 'package:googleapis_auth/auth.dart';

import 'utils.dart';

const _sheetsEndpoint = 'https://sheets.googleapis.com/v4/spreadsheets/';
const _filesEndpoint = 'https://www.googleapis.com/drive/v2/files/';

/// Manages googleapis auth and [Spreadsheet] fetching.
class GSheets {
  AutoRefreshingAuthClient _client;
  final ServiceAccountCredentials _credentials;
  final List<String> _scopes;

  /// Creates a new [GSheets].
  ///
  /// [credentials] must be provided.
  ///
  /// [scopes] defaults to [SheetsApi.SpreadsheetsScope].
  GSheets({
    @required ServiceAccountCredentials credentials,
    List<String> scopes = const [v4.SheetsApi.SpreadsheetsScope],
  })  : _credentials = credentials,
        _scopes = scopes;

  /// Returns Future [AutoRefreshingAuthClient].
  Future<AutoRefreshingAuthClient> get client async {
    if (_client == null) {
      _client = await clientViaServiceAccount(_credentials, _scopes);
    }
    return _client;
  }

  /// Fetches and returns Future [Spreadsheet].
  Future<Spreadsheet> fetch(String spreadsheetId) async {
    final client = await this.client;
    final response = await client.get('$_sheetsEndpoint$spreadsheetId');

    final sheets = (jsonDecode(response.body)['sheets'] as List)
        .map((sheetJson) =>
            Worksheet._fromSheetJson(sheetJson, client, spreadsheetId))
        .toList();

    return Spreadsheet._(
      client,
      spreadsheetId,
      sheets,
    );
  }

  static Future<Response> _batchUpdate(
    AutoRefreshingAuthClient client,
    String spreadsheetId,
    List<Map<String, dynamic>> requests,
  ) =>
      client.post(
        '$_sheetsEndpoint$spreadsheetId:batchUpdate',
        body: jsonEncode({'requests': requests}),
      );
}

/// Representation of a [Spreadsheet], manages [Worksheet]s.
class Spreadsheet {
  final AutoRefreshingAuthClient _client;
  final String id;
  final List<Worksheet> sheets;

  Spreadsheet._(this._client, this.id, this.sheets);

  /// Refreshes [Spreadsheet].
  ///
  /// Should be called if you believe, that spreadsheet has been changed
  /// by another user (such as added/deleted/renamed worksheets).
  ///
  /// Returns Future `true` in case of success.
  Future<bool> refresh() async {
    final response = await _client.get(
      '$_sheetsEndpoint$id',
    );
    if (response.statusCode == 200) {
      final sheets = (jsonDecode(response.body)['sheets'] as List)
          .map((sheetJson) => Worksheet._fromSheetJson(sheetJson, _client, id))
          .toList();
      this.sheets.clear();
      this.sheets.addAll(sheets);
      return true;
    }
    return false;
  }

  /// Returns [Worksheet] with [title].
  ///
  /// Returns `null` if [Worksheet] with [title] not found.
  Worksheet getWorksheetByTitle(String title) {
    return sheets.firstWhere(
      (sheet) => sheet._title == title,
      orElse: () => null,
    );
  }

  /// Returns [Worksheet] with [id].
  ///
  /// Returns `null` if [Worksheet] with [id] not found.
  Worksheet getWorksheetById(int id) {
    return sheets.firstWhere(
      (sheet) => sheet.id == id,
      orElse: () => null,
    );
  }

  /// Returns [Worksheet] with [index].
  ///
  /// Returns `null` if [Worksheet] with [index] not found.
  Worksheet getWorksheetByIndex(int index) {
    return sheets.firstWhere(
      (sheet) => sheet.index == index,
      orElse: () => null,
    );
  }

  /// Adds new [Worksheet] with specified [title], [rows] and [columns].
  ///
  /// [title] - title of a new [Worksheet]
  /// [rows] - row count of a new [Worksheet]
  /// [columns] - column count of a new [Worksheet]
  ///
  /// Returns Future of created [Worksheet].
  ///
  /// Throws [GSheetsException] if sheet with [title] already exists, or
  /// if [rows] or [columns] value is invalid.
  Future<Worksheet> addWorksheet(String title, int rows, int columns) async {
    checkCR(columns, rows);
    final response = await GSheets._batchUpdate(_client, id, [
      {
        'addSheet': {
          'properties': {
            'title': title,
            'sheetType': 'GRID',
            'gridProperties': {
              'rowCount': rows,
              'columnCount': columns,
            }
          },
        }
      }
    ]);
    checkResponse(response);
    final addSheetJson = (jsonDecode(response.body)['replies'] as List)?.first;
    if (addSheetJson != null) {
      sheets
          .add(Worksheet._fromSheetJson(addSheetJson['addSheet'], _client, id));
    }
    return getWorksheetByTitle(title);
  }

  /// Copies [ws] with specified [title] and [index].
  ///
  /// Returns Future of created [Worksheet].
  ///
  /// Throws [GSheetsException] if sheet with [title] already exists.
  Future<Worksheet> copyWorksheet(
    Worksheet ws,
    String title, [
    int index,
  ]) async {
    final response = await GSheets._batchUpdate(_client, id, [
      {
        'duplicateSheet': {
          'sourceSheetId': ws.id,
          'insertSheetIndex': index,
          'newSheetId': null,
          'newSheetName': title
        }
      }
    ]);
    checkResponse(response);
    final duplicateSheetJson =
        (jsonDecode(response.body)['replies'] as List)?.first;
    if (duplicateSheetJson != null) {
      sheets.add(Worksheet._fromSheetJson(
          duplicateSheetJson['duplicateSheet'], _client, id));
    }
    return getWorksheetByTitle(title);
  }

  /// Deletes [ws].
  ///
  /// Returns `true` in case of success.
  ///
  /// Throws [GSheetsException] if something goes wrong.
  Future<bool> deleteWorksheet(Worksheet ws) async {
    final response = await GSheets._batchUpdate(_client, id, [
      {
        'deleteSheet': {'sheetId': ws.id}
      }
    ]);
    checkResponse(response);
    sheets.remove(ws);
    return true;
  }

  /// Returns Future list of [Permission].
  ///
  /// Requires [SheetsApi.DriveScope]
  ///
  /// Throws Exception if auth does not include DriveScope
  /// Throws GSheetsException if DriveScope is not configured
  Future<List<Permission>> getPermissions() async {
    final response = await _client.get('$_filesEndpoint$id/permissions');
    checkResponse(response);
    return (jsonDecode(response.body)['items'] as List)
            ?.map((json) => Permission._fromJson(json))
            ?.toList() ??
        [];
  }

  /// Returns Future [Permission] by email.
  ///
  /// Requires [SheetsApi.DriveScope]
  ///
  /// [email] email of requested permission
  ///
  /// Returns `null` if [email] not found.
  ///
  /// Throws Exception if auth does not include DriveScope
  /// Throws GSheetsException if DriveScope is not configured
  Future<Permission> getPermission(String email) async {
    final response = await _client.get('$_filesEndpoint$id/permissions');
    checkResponse(response);
    return (jsonDecode(response.body)['items'] as List)
        ?.map((json) => Permission._fromJson(json))
        ?.firstWhere((it) => it.email == email, orElse: () => null);
  }

  /// Shares [Spreadsheet].
  ///
  /// Requires [SheetsApi.DriveScope]
  ///
  /// [user] - the email address or domain name for the entity
  /// [type] - the account type
  /// [role] - the primary role for this user
  /// [withLink] - whether the link is required for this permission
  ///
  /// Returns Future of shared [Permission].
  ///
  /// Throws Exception if auth does not include DriveScope
  /// Throws GSheetsException if DriveScope is not configured
  Future<Permission> share({
    @required String user,
    PermType type = PermType.user,
    PermRole role = PermRole.reader,
    bool withLink = false,
  }) async {
    final response = await _client.post(
      '$_filesEndpoint$id/permissions',
      body: jsonEncode({
        'value': user,
        'type': Permission._parseType(type),
        'role': Permission._parseRole(role),
        'withLink': withLink,
      }),
      headers: {'Content-type': 'application/json'},
    );
    checkResponse(response);
    return Permission._fromJson(jsonDecode(response.body));
  }

  /// Revokes permission by [id].
  ///
  /// [id] - permission id to remove
  ///
  /// Returns `true` in case of success
  ///
  /// Throws Exception if auth does not include DriveScope
  /// Throws GSheetsException if DriveScope is not configured
  Future<bool> revokePermissionById(String id) async {
    final response = await _client.delete(
      '$_filesEndpoint${this.id}/permissions/$id',
    );
    return response.statusCode == 204;
  }

  /// Revokes permission by [email].
  ///
  /// [email] - email to remove permission for
  ///
  /// Returns `true` in case of success
  ///
  /// Throws Exception if auth does not include DriveScope
  /// Throws GSheetsException if DriveScope is not configured
  Future<bool> revokePermissionByEmail(String email) async {
    final permission = await getPermission(email);
    if (permission == null) {
      return false;
    }
    return revokePermissionById(permission.id);
  }
}

/// Main dimensions for [Worksheet]
enum SheetDimen { column, row }

/// Permission types
enum PermType { user, group, domain, any }

/// Permission roles
enum PermRole { owner, writer, reader }

/// Representation of a permission.
class Permission {
  final String id;
  final String name;
  final String email;
  final String type;
  final String role;
  final bool deleted;

  static const _typeUser = 'user';
  static const _typeGroup = 'group';
  static const _typeDomain = 'domain';
  static const _typeAny = 'anyone';
  static const _roleOwner = 'owner';
  static const _roleWriter = 'writer';
  static const _roleReader = 'reader';

  Permission._({
    this.id,
    this.name,
    this.email,
    this.type,
    this.role,
    this.deleted,
  });

  static String _parseType(PermType type) {
    switch (type) {
      case PermType.user:
        return _typeUser;
      case PermType.group:
        return _typeGroup;
      case PermType.domain:
        return _typeDomain;
      case PermType.any:
        return _typeAny;
    }
    return null;
  }

  static String _parseRole(PermRole role) {
    switch (role) {
      case PermRole.owner:
        return _roleOwner;
      case PermRole.writer:
        return _roleWriter;
      case PermRole.reader:
        return _roleReader;
    }
    return null;
  }

  factory Permission._fromJson(Map<String, dynamic> json) {
    return Permission._(
      id: json['id'] as String,
      name: json['name'] as String,
      email: json['emailAddress'] as String,
      type: json['type'] as String,
      role: json['role'] as String,
      deleted: json['deleted'] as bool,
    );
  }

  @override
  String toString() => '''
  id: $id
  name: $name
  email: $email
  type: $type
  role: $role
  deleted: $deleted''';
}

/// Representation of a [Worksheet].
class Worksheet {
  final AutoRefreshingAuthClient _client;
  final String spreadsheetId;
  final int id;
  String _title;
  final int index;
  _WorksheetAsCells _cells;
  _WorksheetAsValues _values;
  int _rowCount;
  int _columnCount;

  /// Current count of available rows.
  int get rowCount => _rowCount;

  /// Current count of available columns.
  int get columnCount => _columnCount;

  String get title => _title;

  /// Interactor for working with [Worksheet] cells as [String] values.
  _WorksheetAsValues get values {
    if (_values == null) {
      _values = _WorksheetAsValues(this);
    }
    return _values;
  }

  /// Interactor for working with [Worksheet] cells as [Cell] objects.
  _WorksheetAsCells get cells {
    if (_cells == null) {
      _cells = _WorksheetAsCells(this);
    }
    return _cells;
  }

  Worksheet._(
    this._client,
    this.spreadsheetId,
    this.id,
    this._title,
    this.index, [
    this._rowCount,
    this._columnCount,
  ]);

  factory Worksheet._fromSheetJson(Map<String, dynamic> sheetJson,
      AutoRefreshingAuthClient client, String sheetsId) {
    return Worksheet._(
      client,
      sheetsId,
      sheetJson['properties']['sheetId'],
      sheetJson['properties']['title'],
      sheetJson['properties']['index'],
      sheetJson['properties']['gridProperties']['rowCount'],
      sheetJson['properties']['gridProperties']['columnCount'],
    );
  }

  /// Updates title of this [Worksheet].
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException] if something goes wrong.
  Future<bool> updateTitle(String title) async {
    if (_title == title || isNullOrEmpty(title)) {
      return false;
    }
    final response = await GSheets._batchUpdate(_client, spreadsheetId, [
      {
        'updateSheetProperties': {
          'properties': {
            'sheetId': id,
            'title': title,
          },
          'fields': 'title',
        }
      }
    ]);
    checkResponse(response);
    _title = title;
    return true;
  }

  /// Deletes columns/rows from [Worksheet].
  ///
  /// [dimen] - deleting dimension
  ///
  /// [index] - index of a column/row to delete
  /// columns start at index 1 (column A), rows start at index 1
  ///
  /// [length] - optional (defaults to 1), the number of columns/rows to delete
  /// starting from [index]
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException] if something goes wrong.
  Future<bool> delete(SheetDimen dimen, int index, {int length = 1}) async {
    checkI(index);
    checkL(length);
    final response = await GSheets._batchUpdate(_client, spreadsheetId, [
      {
        'deleteDimension': {
          'range': {
            'sheetId': id,
            "dimension":
                dimen == SheetDimen.column ? DIMEN_COLUMNS : DIMEN_ROWS,
            "startIndex": index - 1,
            "endIndex": index - 1 + length,
          },
        }
      }
    ]);
    checkResponse(response);
    return true;
  }

  /// Inserts new columns/rows to [Worksheet].
  ///
  /// [dimen] - inserting dimension
  ///
  /// [index] - index of a column/row to insert
  /// columns start at index 1 (column A), rows start at index 1
  ///
  /// [length] - optional (defaults to 1), the number of columns/rows to insert
  ///
  /// [inheritFromBefore] - optional (defaults to `false`), if true, tells the
  /// API to give the new columns/rows the same properties as the prior
  /// column/row, otherwise the new columns/rows acquire the properties of
  /// those that follow them,
  /// cannot be true if inserting a column/row with index 1
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException] if something goes wrong.
  Future<bool> insert(
    SheetDimen dimen,
    int index, {
    int length = 1,
    bool inheritFromBefore = false,
  }) async {
    checkC(index);
    checkL(length);
    final response = await GSheets._batchUpdate(_client, spreadsheetId, [
      {
        'insertDimension': {
          'range': {
            'sheetId': id,
            "dimension":
                dimen == SheetDimen.column ? DIMEN_COLUMNS : DIMEN_ROWS,
            "startIndex": index - 1,
            "endIndex": index - 1 + length,
          },
          "inheritFromBefore": inheritFromBefore
        }
      }
    ]);
    checkResponse(response);
    return true;
  }

  /// Moves columns/rows.
  ///
  /// [from] - index of a first column/row to move
  /// columns start at index 1 (column A), rows start at index 1
  ///
  /// [length] - optional (defaults to 1), the number of columns/rows to move
  ///
  /// [to] - new index of a last column/row moved
  /// must be in a range from [from] to [from] + [length]
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException] if something goes wrong.
  Future<bool> move(
    SheetDimen dimen, {
    @required int from,
    @required int to,
    int length = 1,
  }) async {
    except(from == to, 'cannot move from $from to $to');
    except(from < 1, 'invalid from ($from)');
    except(to < 1, 'invalid to ($to)');
    checkL(length);
    // correct values for from > to
    final cFrom = from < to ? from : to;
    final cTo = from < to ? to : from + length - 1;
    final cLength = from < to ? length : from - to;
    final response = await GSheets._batchUpdate(_client, spreadsheetId, [
      {
        'moveDimension': {
          'source': {
            'sheetId': id,
            "dimension":
                dimen == SheetDimen.column ? DIMEN_COLUMNS : DIMEN_ROWS,
            "startIndex": cFrom - 1,
            "endIndex": cFrom - 1 + cLength,
          },
          "destinationIndex": cTo
        }
      }
    ]);
    checkResponse(response);
    return true;
  }

  Future<List<String>> _get(String range, String dimension) async {
    final response = await _client.get(
      '$_sheetsEndpoint$spreadsheetId/values/$range?majorDimension=$dimension',
    );
    return ((jsonDecode(response.body)['values'] as List)?.first as List)
            ?.cast<String>() ??
        [];
  }

  Future<List<List<String>>> _getAll(String range, String dimension) async {
    final response = await _client.get(
      '$_sheetsEndpoint$spreadsheetId/values/$range?majorDimension=$dimension',
    );
    final list = <List<String>>[];
    (jsonDecode(response.body)['values'] as List)?.forEach((sublist) {
      list.add(sublist?.cast<String>() ?? <String>[]);
    });
    return list;
  }

  Future<bool> _update({
    List<String> values,
    String majorDimension,
    String range,
  }) async {
    final response = await _client.put(
      '$_sheetsEndpoint$spreadsheetId/values/$range?valueInputOption=USER_ENTERED',
      body: jsonEncode(
        {
          "range": range,
          "majorDimension": majorDimension,
          "values": [values],
        },
      ),
    );
    return response.statusCode == 200;
  }

  Future<bool> _clear(String range) async {
    final response = await _client.post(
      '$_sheetsEndpoint$spreadsheetId/values/$range:clear',
    );
    print(response.body);
    return response.statusCode == 200;
  }

  Future<String> _rowRange(int row, int column, [int length = -1]) async {
    await _expand(row, column + length - 1);
    String label = getColumnLetter(column);
    String labelTo = length > 0 ? getColumnLetter(column + length - 1) : '';
    return "'$_title'!${label}${row}:${labelTo}${row}";
  }

  Future<String> _columnRange(int column, int row, [int length = -1]) async {
    await _expand(row + length - 1, column);
    String label = getColumnLetter(column);
    String to = length > 0 ? '${row + length - 1}' : '';
    return "'$_title'!$label${row}:$label$to";
  }

  Future<String> _allColumnsRange(
    int column,
    int row, [
    int length = -1,
  ]) async {
    await _expand(row + length - 1, column);
    String fromLabel = getColumnLetter(column);
    String toLabel = getColumnLetter(columnCount);
    int to = length > 0 ? row + length - 1 : rowCount;
    return "'$_title'!$fromLabel${row}:$toLabel$to";
  }

  Future<String> _allRowsRange(
    int row,
    int column, [
    int length = -1,
  ]) async {
    await _expand(row, column + length - 1);
    String label = getColumnLetter(column);
    String toLabel = length > 0
        ? getColumnLetter(column + length - 1)
        : getColumnLetter(columnCount);
    return "'$_title'!${label}${row}:$toLabel$rowCount";
  }

  Future<void> _expand(int rows, int cols) async {
    bool changed = false;
    if (_rowCount < rows) {
      _rowCount = rows;
      changed = true;
    }
    if (_columnCount < cols) {
      _columnCount = cols;
      changed = true;
    }
    if (changed) {
      await GSheets._batchUpdate(_client, spreadsheetId, [
        {
          'updateSheetProperties': {
            'properties': {
              'sheetId': id,
              'gridProperties': {
                'rowCount': _rowCount,
                'columnCount': _columnCount,
              }
            },
            'fields': 'gridProperties/rowCount,gridProperties/columnCount'
          }
        }
      ]);
    }
  }

  /// Clears the whole [Worksheet].
  ///
  /// Returns Future `true` in case of success.
  Future<bool> clear() => _clear("'$title'");

  /// Clears specified column.
  ///
  /// [column] - index of a column to clear
  /// columns start at index 1 (column A)
  ///
  /// [row] - optional (defaults to 1), index of a row that the column
  /// will be cleared from (values before [row] will remain uncleared)
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), number of cells to clear in the
  /// column
  ///
  /// Returns Future `true` in case of success.
  Future<bool> clearColumn(
    int column, {
    int row = 1,
    int length = -1,
  }) async {
    checkCR(column, row);
    return _clear(await _columnRange(column, row, length));
  }

  /// Clears specified row.
  ///
  /// [row] - index of a row to clear
  /// rows start at index 1
  ///
  /// [column] - optional (defaults to 1), index of a column that the row
  /// will be cleared from (values before [column] will remain uncleared)
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), number of cells to clear in the row
  ///
  /// Returns Future `true` in case of success.
  Future<bool> clearRow(
    int row, {
    int column = 1,
    int length = -1,
  }) async {
    checkCR(column, row);
    return _clear(await _rowRange(row, column, length));
  }
}

class _WorksheetAsValues {
  final Worksheet _ws;
  _ValueMapper _map;

  _WorksheetAsValues(this._ws);

  _ValueMapper get map {
    if (_map == null) {
      _map = _ValueMapper(this);
    }
    return _map;
  }

  /// Fetches specified column.
  ///
  /// [column] - index of a requested column
  /// columns start at index 1 (column A)
  ///
  /// [row] - optional (defaults to 1), index of a row that requested column
  /// starts from (values before [row] will be skipped)
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all values starting from [row] will be returned
  ///
  /// Returns column as Future [List] of [String].
  Future<List<String>> column(
    int column, {
    int row = 1,
    int length = -1,
  }) async {
    checkCR(column, row);
    final range = await _ws._columnRange(column, row, length);
    return _ws._get(range, DIMEN_COLUMNS);
  }

  /// Fetches specified row.
  ///
  /// [row] - index of a requested row
  /// rows start at index 1
  ///
  /// [column] - optional (defaults to 1), index of a column that requested row
  /// starts from (values before [column] will be skipped)
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all values starting from [column] will be returned
  ///
  /// Returns row as Future [List] of [String].
  Future<List<String>> row(
    int row, {
    int column = 1,
    int length = -1,
  }) async {
    checkCR(column, row);
    final range = await _ws._rowRange(row, column, length);
    return _ws._get(range, DIMEN_ROWS);
  }

  /// Fetches column by its name.
  ///
  /// [key] - name of a requested column
  /// The first row considered to be column names
  ///
  /// [row] - optional (defaults to 2), index of a row that requested column
  /// starts from (values before [row] will be skipped)
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all values starting from [row] will be returned
  ///
  /// Returns column as Future [List] of [String].
  Future<List<String>> columnByKey(
    String key, {
    int row = 2,
    int length = -1,
  }) async {
    final column = await columnIndexOf(key, add: false);
    if (column < 1) return null;
    return this.column(column, row: row, length: length);
  }

  /// Fetches row by its name.
  ///
  /// [key] - name of a requested row
  /// The column A considered to be row names
  ///
  /// [column] - optional (defaults to 2), index of a column that requested row
  /// starts from (values before [column] will be skipped)
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all values starting from [column] will be returned
  ///
  /// Returns row as Future [List] of [String].
  Future<List<String>> rowByKey(
    String key, {
    int column = 2,
    int length = -1,
  }) async {
    final row = await rowIndexOf(key, add: false);
    if (row < 1) return null;
    return this.row(row, column: column, length: length);
  }

  /// Fetches last column.
  ///
  /// [row] - optional (defaults to 1), index of a row that requested column
  /// starts from (values before [row] will be skipped)
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all values starting from [row] will be returned
  ///
  /// Returns last column as Future [List] of [String].
  Future<List<String>> lastColumn({
    int row = 1,
    int length = -1,
  }) async {
    final column = (await this.row(1)).length;
    return this.column(column, row: row, length: length);
  }

  /// Fetches last row.
  ///
  /// [column] - optional (defaults to 1), index of a column that requested row
  /// starts from (values before [column] will be skipped)
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all values starting from [column] will be returned
  ///
  /// Returns last row as Future [List] of [String].
  Future<List<String>> lastRow({
    int column = 1,
    int length = -1,
  }) async {
    final row = (await this.column(1)).length;
    return this.row(row, column: column, length: length);
  }

  /// Fetches all columns.
  ///
  /// [column] - optional (defaults to 1), index of a first returned column
  /// (columns before [column] will be skipped)
  /// columns start at index 1 (column A)
  ///
  /// [row] - optional (defaults to 1), index of a row that columns start from
  /// (values before [row] will be skipped)
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested columns
  /// if length is `-1`, all values starting from [row] will be returned
  ///
  /// Returns all columns as Future [List] of [List].
  Future<List<List<String>>> allColumns({
    int column = 1,
    int row = 1,
    int length = -1,
  }) async {
    checkCR(column, row);
    final range = await _ws._allColumnsRange(column, row, length);
    return _ws._getAll(range, DIMEN_COLUMNS);
  }

  /// Fetches all rows.
  ///
  /// [row] - optional (defaults to 1), index of a first returned row
  /// (rows before [row] will be skipped)
  /// rows start at index 1
  ///
  /// [column] - optional (defaults to 1), index of a column that rows start
  /// from (values before [column] will be skipped)
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested rows
  /// if length is `-1`, all values starting from [column] will be returned
  ///
  /// Returns all rows as Future [List] of [List].
  Future<List<List<String>>> allRows({
    int row = 1,
    int column = 1,
    int length = -1,
  }) async {
    checkCR(column, row);
    final range = await _ws._allRowsRange(row, column, length);
    return _ws._getAll(range, DIMEN_ROWS);
  }

  /// Fetches cell's value.
  ///
  /// [column] - column index of a requested cell's value
  /// columns start at index 1 (column A)
  ///
  /// [row] - row index of a requested cell's value
  /// rows start at index 1
  ///
  /// Returns cell's value as Future [String].
  Future<String> value({
    @required int column,
    @required int row,
  }) async {
    checkCR(column, row);
    final range = await _ws._columnRange(column, row);
    return getOrEmpty(await _ws._get(range, DIMEN_COLUMNS));
  }

  /// Fetches cell's value by names of its column and row.
  ///
  /// [rowKey] - name of a row with requested cell's value
  /// The column A considered to be row names
  ///
  /// [columnKey] - name of a column with requested cell's value
  /// The first row considered to be column names
  ///
  /// Returns cell's value as Future [String].
  ///
  /// Returns `null` if either [rowKey] or [columnKey] not found.
  Future<String> valueByKeys({
    @required String rowKey,
    @required String columnKey,
  }) async {
    final column = await columnIndexOf(columnKey, add: false);
    if (column < 1) return null;
    final row = await rowIndexOf(rowKey, add: false);
    if (row < 1) return null;
    final range = await _ws._columnRange(column, row);
    return getOrEmpty(await _ws._get(range, DIMEN_COLUMNS));
  }

  /// Updates cell's value to [value].
  ///
  /// [row] - row index to insert [value] to
  /// rows start at index 1
  ///
  /// [column] - column index to insert [value] to
  /// columns start at index 1 (column A)
  ///
  /// Returns Future `true` in case of success.
  Future<bool> insertValue({
    @required int column,
    @required int row,
    @required String value,
  }) async {
    checkCR(column, row);
    return _ws._update(
      values: [value ?? ''],
      range: await _ws._columnRange(column, row),
      majorDimension: DIMEN_COLUMNS,
    );
  }

  /// Updates cell's value to [value] by names of its column and row.
  ///
  /// [columnKey] - name of a column to insert [value] to
  /// The first row considered to be column names
  ///
  /// [rowKey] - name of a row to insert [value] to
  /// The column A considered to be row names
  ///
  /// Returns Future `true` in case of success.
  Future<bool> insertValueByKeys({
    @required String columnKey,
    @required String rowKey,
    @required String value,
  }) async {
    int column = await columnIndexOf(columnKey, add: true);
    int row = await rowIndexOf(rowKey, add: true);
    return insertValue(column: column, row: row, value: value);
  }

  /// Returns index of a column with [key] value in [row].
  ///
  /// [key] - value to look for
  ///
  /// [row] - optional (defaults to 1), row index in which [key] is looked for
  /// rows start at index 1
  ///
  /// [add] - optional (defaults to `false`), whether the [key] should be added
  /// to [row] in case of absence
  ///
  /// Returns Future `-1` if not found and [add] is false.
  Future<int> columnIndexOf(String key, {bool add = true, int row = 1}) async {
    except(isNullOrEmpty(key), 'invalid key ($key)');
    final columnKeys = await this.row(row);
    int column = columnKeys.indexOf(key) + 1;
    if (column < 1) {
      column = -1;
      if (add) {
        column = columnKeys.length + 1;
        final isAdded = await _ws._update(
          values: [key],
          range: await _ws._columnRange(column, row),
          majorDimension: DIMEN_COLUMNS,
        );
        if (!isAdded) column = -1;
      }
    }
    return column;
  }

  /// Returns index of a row with [key] value in [column].
  ///
  /// [key] - value to look for
  ///
  /// [column] - optional (defaults to 1), column index in which [key] is
  /// looked for
  /// columns start at index 1 (column A)
  ///
  /// [add] - optional (defaults to `false`), whether the [key] should be added
  /// to [column] in case of absence
  ///
  /// Returns Future `-1` if [key] is not found and [add] is false.
  Future<int> rowIndexOf(String key, {bool add = false, column = 1}) async {
    except(isNullOrEmpty(key), 'invalid key ($key)');
    final rowKeys = await this.column(column);
    int row = rowKeys.indexOf(key) + 1;
    if (row < 1) {
      row = -1;
      if (add) {
        row = rowKeys.length + 1;
        final isAdded = await _ws._update(
          values: [key],
          range: await _ws._columnRange(column, row),
          majorDimension: DIMEN_COLUMNS,
        );
        if (!isAdded) row = -1;
      }
    }
    return row;
  }

  /// Updates column values with [values].
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [column] - column index to insert [values] to
  /// columns start at index 1 (column A)
  ///
  /// [row] - optional (defaults to 1), row index for the first inserted value
  /// of [values]
  /// rows start at index 1
  ///
  /// Returns Future `true` in case of success.
  Future<bool> insertColumn({
    @required List<String> values,
    @required int column,
    int row = 1,
  }) async {
    checkCR(column, row);
    checkV(values);
    return _ws._update(
      values: values,
      range: await _ws._columnRange(column, row, values.length),
      majorDimension: DIMEN_COLUMNS,
    );
  }

  /// Updates column by its name values with [values].
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [key] - name of a column to insert [values] to
  /// The first row considered to be column names
  ///
  /// [row] - optional (defaults to 2), row index for the first inserted value
  /// of [values]
  /// rows start at index 1
  ///
  /// Returns Future `true` in case of success.
  Future<bool> insertColumnByKey({
    @required List<String> values,
    @required String key,
    int row = 2,
  }) async {
    final column = await columnIndexOf(key, add: true);
    return insertColumn(values: values, row: row, column: column);
  }

  /// Appends column.
  ///
  /// The column index that [values] will be inserted to is the column index
  /// after last cell in the first row
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [row] - optional (defaults to 1), row index for the first inserted value
  /// of [values]
  /// rows start at index 1
  ///
  /// Returns Future `true` in case of success.
  Future<bool> appendColumn({
    @required List<String> values,
    int row = 1,
  }) async {
    final column = (await this.row(1)).length + 1;
    return insertColumn(values: values, column: column, row: row);
  }

  /// Updates row values with [values].
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [row] - row index to insert [values] to
  /// rows start at index 1
  ///
  /// [column] - optional (defaults to 1), column index for the first inserted
  /// value of [values]
  /// columns start at index 1 (column A)
  ///
  /// Returns Future `true` in case of success.
  Future<bool> insertRow({
    @required List<String> values,
    @required int row,
    int column = 1,
  }) async {
    checkCR(column, row);
    checkV(values);
    return _ws._update(
      values: values,
      range: await _ws._rowRange(row, column, values.length),
      majorDimension: DIMEN_ROWS,
    );
  }

  /// Updates row by its name values with [values].
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [key] - name of a row to insert [values] to
  /// The column A considered to be row names
  ///
  /// [column] - optional (defaults to 2), column index for the first inserted
  /// value of [values]
  /// columns start at index 1 (column A)
  ///
  /// Returns Future `true` in case of success.
  Future<bool> insertRowByKey({
    @required List<String> values,
    @required String key,
    int column = 2,
  }) async {
    int row = await rowIndexOf(key, add: true);
    return insertRow(values: values, column: column, row: row);
  }

  /// Appends row.
  ///
  /// The row index that [values] will be inserted to is the row index after
  /// last cell in the column A
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [column] - optional (defaults to 1), column index for the first inserted
  /// value of [values]
  /// columns start at index 1 (column A)
  ///
  /// Returns Future `true` in case of success.
  Future<bool> appendRow({
    @required List<String> values,
    int column = 1,
  }) async {
    final row = (await this.column(1)).length + 1;
    return insertRow(values: values, column: column, row: row);
  }
}

class _ValueMapper {
  final _WorksheetAsValues _values;

  _ValueMapper(this._values);

  /// Fetches specified column, maps it to other column and returns map.
  ///
  /// [column] - index of a requested column (values of returned map)
  /// columns start at index 1 (column A)
  ///
  /// [row] - optional (defaults to 1), index of a row that requested column
  /// starts from (values before [row] will be skipped)
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all values starting from [row] will be returned
  ///
  /// [mapTo] - optional (defaults to 1), index of a column to map values to
  /// (keys of returned map)
  /// columns start at index 1 (column A)
  ///
  /// Returns column as Future [Map] of [String] to [String].
  Future<Map<String, String>> column(
    int column, {
    int row = 1,
    int length = -1,
    int mapTo = 1,
  }) async {
    checkM(column, mapTo);
    final values = await _values.column(column, row: row, length: length);
    final keys = await _values.column(mapTo, row: row, length: length);
    final map = <String, String>{};
    mapKeysToValues(keys, values, map, '', null);
    return map;
  }

  /// Fetches specified row, maps it to other row and returns map.
  ///
  /// [row] - index of a requested row (values of returned map)
  /// rows start at index 1
  ///
  /// [column] - optional (defaults to 1), index of a column that requested row
  /// starts from (values before [column] will be skipped)
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all values starting from [column] will be returned
  ///
  /// [mapTo] - optional (defaults to 1), index of a row to map values to
  /// (keys of returned map)
  /// rows start at index 1
  ///
  /// Returns row as Future [Map] of [String] to [String].
  Future<Map<String, String>> row(
    int row, {
    int column = 1,
    int length = -1,
    int mapTo = 1,
  }) async {
    checkM(row, mapTo);
    final values = await _values.row(row, column: column, length: length);
    final keys = await _values.row(mapTo, column: column, length: length);
    final map = <String, String>{};
    mapKeysToValues(keys, values, map, '', null);
    return map;
  }

  /// Fetches column by its name, maps it to other column and returns map.
  ///
  /// The first row considered to be column names
  ///
  /// [key] - name of a requested column (values of returned map)
  ///
  /// [row] - optional (defaults to 1), index of a row that requested column
  /// starts from (values before [row] will be skipped)
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all values starting from [row] will be returned
  ///
  /// [mapTo] - optional, name of a column to map values to (keys of returned
  /// map), if [mapTo] is `null` then values will be mapped to column A
  /// columns start at index 1 (column A)
  ///
  /// Returns column as Future [Map] of [String] to [String].
  Future<Map<String, String>> columnByKey(
    String key, {
    int row = 2,
    int length = -1,
    String mapTo,
  }) async {
    checkM(key, mapTo);
    final column = await _values.columnIndexOf(key, add: false);
    if (column < 1) return null;
    if (isNullOrEmpty(mapTo)) {
      return this.column(column, row: row, length: length);
    }
    final mapToIndex = await _values.columnIndexOf(mapTo, add: false);
    if (mapToIndex < 1) return null;
    return this.column(column, row: row, length: length, mapTo: mapToIndex);
  }

  /// Fetches row by its name, maps it to other row, and returns map.
  ///
  /// The column A considered to be row names
  ///
  /// [key] - name of a requested row (values of returned map)
  ///
  /// [column] - optional (defaults to 1), index of a column that requested row
  /// starts from (values before [column] will be skipped)
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all values starting from [column] will be returned
  ///
  /// [mapTo] - optional, name of a row to map values to (keys of returned
  /// map), if [mapTo] is `null` then values will be mapped to first row
  /// rows start at index 1
  ///
  /// Returns row as Future [Map] of [String] to [String].
  Future<Map<String, String>> rowByKey(
    String key, {
    int column = 2,
    int length = -1,
    String mapTo,
  }) async {
    checkM(key, mapTo);
    final row = await _values.rowIndexOf(key, add: false);
    if (row < 1) return null;
    if (isNullOrEmpty(mapTo)) {
      return this.row(row, column: column, length: length);
    }
    final mapToIndex = await _values.rowIndexOf(mapTo, add: false);
    if (mapToIndex < 1) return null;
    return this.row(row, column: column, length: length, mapTo: mapToIndex);
  }

  /// Fetches last column, maps it to other column and returns map.
  ///
  /// [row] - optional (defaults to 1), index of a row that requested column
  /// starts from (values before [row] will be skipped)
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all values starting from [row] will be returned
  ///
  /// [mapTo] - optional (defaults to 1), index of a column to map values to
  /// (keys of returned map)
  /// columns start at index 1 (column A)
  ///
  /// Returns column as Future [Map] of [String] to [String].
  Future<Map<String, String>> lastColumn({
    int row = 1,
    int length = -1,
    int mapTo = 1,
  }) async {
    final column = (await _values.row(1)).length;
    if (column < 1) return null;
    return this.column(column, row: row, length: length, mapTo: mapTo);
  }

  /// Fetches last row, maps it to other row and returns map.
  ///
  /// [column] - optional (defaults to 1), index of a column that requested row
  /// starts from (values before [column] will be skipped)
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all values starting from [column] will be returned
  ///
  /// [mapTo] - optional (defaults to 1), index of a row to map values to
  /// (keys of returned map)
  /// rows start at index 1
  ///
  /// Returns row as Future [Map] of [String] to [String].
  Future<Map<String, String>> lastRow({
    int column = 1,
    int length = -1,
    int mapTo = 1,
  }) async {
    final row = (await _values.column(1)).length;
    if (row < 1) return null;
    return this.row(row, column: column, length: length, mapTo: mapTo);
  }

  /// Updates column values with values from [map].
  ///
  /// [map] - map containing values to insert (not null nor empty)
  ///
  /// [column] - column index to insert values of [map] to,
  /// columns start at index 1 (column A)
  ///
  /// [row] - optional (defaults to 1), row index for the first inserted value
  /// rows start at index 1
  ///
  /// [mapTo] - optional (defaults to 1), index of a column to which
  /// keys of the [map] will be mapped to
  /// columns start at index 1 (column A)
  ///
  /// [appendMissing] - optional (defaults to `false`), whether keys of [map]
  /// (with its related values) that are not present in a [mapTo] column
  /// should be added
  ///
  /// [overwrite] - optional (defaults to `false`), whether clear cells of
  /// [column] if [map] does not contain value for them
  ///
  /// Returns Future `true` in case of success.
  Future<bool> insertColumn({
    @required Map<String, String> map,
    @required int column,
    int row = 1,
    int mapTo = 1,
    bool appendMissing = false,
    bool overwrite = false,
  }) async {
    checkCR(column, row);
    checkM(column, mapTo);
    checkV(map);
    final columnMap = Map()..addAll(map);
    final rows = await _values.column(mapTo, row: row);

    final newColumn = <String>[];
    if (overwrite) {
      for (String row in rows) {
        newColumn.add(columnMap.remove(row) ?? '');
      }
    } else {
      for (String row in rows) {
        newColumn.add(columnMap.remove(row));
      }
    }

    if (appendMissing && columnMap.isNotEmpty) {
      final newKeys = <String>[];
      for (MapEntry entry in columnMap.entries) {
        newKeys.add(entry.key);
        newColumn.add(entry.value);
      }
      await _values.insertColumn(
        row: row + rows.length,
        column: mapTo,
        values: newKeys,
      );
    }

    return _values.insertColumn(row: row, column: column, values: newColumn);
  }

  /// Updates column values with values from [map] by column names.
  ///
  /// The first row considered to be column names
  ///
  /// [map] - map containing values to insert (not null nor empty)
  ///
  /// [key] - name of a requested column (values of returned map)
  ///
  /// [row] - optional (defaults to 2), row index for the first inserted value
  /// rows start at index 1
  ///
  /// [mapTo] - optional, name of a column to which keys of the [map] will be
  /// mapped to, if [mapTo] is `null` then values will be mapped to column A
  ///
  /// [appendMissing] - optional (defaults to `false`), whether keys of [map]
  /// (with its related values) that are not present in a [mapTo] column
  /// should be added
  ///
  /// [overwrite] - optional (defaults to `false`), whether clear cells of
  /// [key] column if [map] does not contain value for them
  ///
  /// Returns Future `true` in case of success.
  Future<bool> insertColumnByKey({
    @required Map<String, String> map,
    @required String key,
    int row = 2,
    String mapTo,
    bool appendMissing = false,
    bool overwrite = false,
  }) async {
    checkM(key, mapTo);
    final column = await _values.columnIndexOf(key, add: true);
    final int mapToIndex = isNullOrEmpty(mapTo)
        ? 1
        : await _values.columnIndexOf(mapTo, add: false);
    return insertColumn(
        map: map,
        column: column,
        row: row,
        mapTo: mapToIndex,
        appendMissing: appendMissing,
        overwrite: overwrite);
  }

  /// Appends column with values from [map].
  ///
  /// [map] - map containing values to insert (not null nor empty)
  ///
  /// [row] - optional (defaults to 1), row index for the first inserted value
  /// rows start at index 1
  ///
  /// [mapTo] - optional (defaults to 1), index of a column to which
  /// keys of the [map] will be mapped to
  /// columns start at index 1 (column A)
  ///
  /// [appendMissing] - optional (defaults to `false`), whether keys of [map]
  /// (with its related values) that are not present in a [mapTo] column
  /// should be added
  ///
  /// Returns Future `true` in case of success.
  Future<bool> appendColumn({
    @required Map<String, String> map,
    int row = 1,
    int mapTo = 1,
    bool appendMissing = false,
  }) async {
    final column = (await _values.row(1)).length + 1;
    if (column < 2) return false;
    return insertColumn(
      map: map,
      column: column,
      row: row,
      mapTo: mapTo,
      appendMissing: appendMissing,
      overwrite: false,
    );
  }

  /// Updates row values with values from [map].
  ///
  /// [map] - map containing values to insert (not null nor empty)
  ///
  /// [row] - row index to insert values of [map] to,
  /// rows start at index 1
  ///
  /// [column] - optional (defaults to 1), column index for the first inserted
  /// value
  /// columns start at index 1 (column A)
  ///
  /// [mapTo] - optional (defaults to 1), index of a row to which
  /// keys of the [map] will be mapped to
  /// rows start at index 1
  ///
  /// [appendMissing] - optional (defaults to `false`), whether keys of [map]
  /// (with its related values) that are not present in a [mapTo] row
  /// should be added
  ///
  /// [overwrite] - optional (defaults to `false`), whether clear cells of
  /// [row] if [map] does not contain value for them
  ///
  /// Returns Future `true` in case of success.
  Future<bool> insertRow({
    @required Map<String, String> map,
    @required int row,
    int column = 1,
    int mapTo = 1,
    bool appendMissing = false,
    bool overwrite = false,
  }) async {
    checkCR(column, row);
    checkM(row, mapTo);
    checkV(map);
    final rowMap = Map()..addAll(map);
    final columns = await _values.row(mapTo, column: column);

    final newRow = <String>[];
    if (overwrite) {
      for (String column in columns) {
        newRow.add(rowMap.remove(column) ?? '');
      }
    } else {
      for (String column in columns) {
        newRow.add(rowMap.remove(column));
      }
    }

    if (appendMissing && rowMap.isNotEmpty) {
      final newKeys = <String>[];
      for (MapEntry entry in rowMap.entries) {
        newKeys.add(entry.key);
        newRow.add(entry.value);
      }
      await _values.insertRow(
        row: mapTo,
        column: column + columns.length,
        values: newKeys,
      );
    }

    return _values.insertRow(row: row, column: column, values: newRow);
  }

  /// Updates row values with values from [map] by column names.
  ///
  /// The column A considered to be row names
  ///
  /// [map] - map containing values to insert (not null nor empty)
  ///
  /// [key] - name of a requested row (values of returned map)
  ///
  /// [column] - optional (defaults to 2), column index for the first inserted
  /// value
  /// columns start at index 1 (column A)
  ///
  /// [mapTo] - optional, name of a column to which keys of the [map] will be
  /// mapped to, if [mapTo] is `null` then values will be mapped to column A
  ///
  /// [appendMissing] - whether keys of [map] (with its related values) that
  /// are not present in a [mapTo] should be inserted into [key] row
  ///
  /// [overwrite] - whether clear cells of [key] row if [map] does not
  /// contain value for them
  ///
  /// Returns Future `true` in case of success.
  Future<bool> insertRowByKey({
    @required Map<String, String> map,
    @required String key,
    int column = 2,
    String mapTo,
    bool appendMissing = false,
    bool overwrite = false,
  }) async {
    checkM(key, mapTo);
    final row = await _values.rowIndexOf(key, add: true);
    final int mapToIndex =
        isNullOrEmpty(mapTo) ? 1 : await _values.rowIndexOf(mapTo, add: false);
    return insertRow(
      map: map,
      column: column,
      row: row,
      mapTo: mapToIndex,
      appendMissing: appendMissing,
      overwrite: overwrite,
    );
  }

  /// Appends row with values from [map].
  ///
  /// [map] - map containing values to insert (not null nor empty)
  ///
  /// [column] - optional (defaults to 1), column index for the first inserted
  /// value
  /// columns start at index 1 (column A)
  ///
  /// [mapTo] - optional (defaults to 1), index of a row to which
  /// keys of the [map] will be mapped to
  /// rows start at index 1
  ///
  /// [appendMissing] - optional (defaults to `false`), whether keys of [map]
  /// (with its related values) that are not present in a [mapTo] row
  /// should be added
  ///
  /// Returns Future `true` in case of success.
  Future<bool> appendRow({
    @required Map<String, String> map,
    int column = 1,
    int mapTo = 1,
    bool appendMissing = false,
  }) async {
    final row = (await _values.column(1)).length + 1;
    if (row < 2) return false;
    return insertRow(
      map: map,
      row: row,
      column: column,
      mapTo: mapTo,
      appendMissing: appendMissing,
      overwrite: false,
    );
  }
}

/// Representation of a [Cell].
class Cell implements Comparable {
  final Worksheet _ws;
  final int rowIndex;
  final int columnIndex;
  String value;

  Cell._(
    this._ws,
    this.rowIndex,
    this.columnIndex,
    this.value,
  );

  /// Returns position of a cell in A1 notation.
  get label => '${getColumnLetter(columnIndex)}${rowIndex}';

  get worksheetTitle => _ws._title;

  /// Updates value of a cell.
  ///
  /// Returns Future `true` in case of success
  Future<bool> post(String value) async {
    this.value = value;
    final range = "'$worksheetTitle'!$label";
    return _ws._update(
      values: [value ?? ''],
      range: range,
      majorDimension: DIMEN_COLUMNS,
    );
  }

  /// Refreshes value of a cell.
  ///
  /// Returns Future `true` if value has been changed.
  Future<bool> refresh() async {
    final before = value;
    value = getOrEmpty(
        await _ws.values.column(columnIndex, row: rowIndex, length: 1));
    return before != value;
  }

  @override
  String toString() => "'$value' at $label";

  @override
  int compareTo(other) {
    return rowIndex + columnIndex - other.rowIndex - other.columnIndex;
  }
}

class _WorksheetAsCells {
  final Worksheet _ws;
  _CellsMapper _map;

  _WorksheetAsCells(this._ws);

  _CellsMapper get map {
    if (_map == null) {
      _map = _CellsMapper(this);
    }
    return _map;
  }

  /// Fetches specified column.
  ///
  /// [column] - index of a requested column
  /// columns start at index 1 (column A)
  ///
  /// [row] - optional (defaults to 1), index of a row that requested column
  /// starts from (cells before [row] will be skipped)
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all cells starting from [row] will be returned
  ///
  /// Returns column as Future [List] of [Cell].
  Future<List<Cell>> column(
    int column, {
    int row = 1,
    int length = -1,
  }) async {
    int index = row;
    return List.unmodifiable(
        (await _ws.values.column(column, row: row, length: length))
            .map((value) {
      return Cell._(
        _ws,
        index++,
        column,
        value,
      );
    }));
  }

  /// Fetches specified row.
  ///
  /// [row] - index of a requested row
  /// rows start at index 1
  ///
  /// [column] - optional (defaults to 1), index of a column that requested row
  /// starts from (cells before [column] will be skipped)
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all cells starting from [column] will be returned
  ///
  /// Returns row as Future [List] of [Cell].
  Future<List<Cell>> row(
    int row, {
    int column = 1,
    int length = -1,
  }) async {
    int index = column;
    return List.unmodifiable(
        (await _ws.values.row(row, column: column, length: length))
            .map((value) {
      return Cell._(
        _ws,
        row,
        index++,
        value,
      );
    }));
  }

  /// Fetches column by its name.
  ///
  /// [key] - name of a requested column
  /// The first row considered to be column names
  ///
  /// [row] - optional (defaults to 2), index of a row that requested column
  /// starts from (cells before [row] will be skipped)
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all cells starting from [row] will be returned
  ///
  /// Returns column as Future [List] of [Cell].
  Future<List<Cell>> columnByKey(
    String key, {
    int row = 2,
    int length = -1,
  }) async {
    final column = await _ws.values.columnIndexOf(key, add: false);
    if (column < 1) return null;
    return this.column(column, row: row, length: length);
  }

  /// Fetches row by its name.
  ///
  /// [key] - name of a requested row
  /// The column A considered to be row names
  ///
  /// [column] - optional (defaults to 2), index of a column that requested row
  /// starts from (cells before [column] will be skipped)
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all cells starting from [column] will be returned
  ///
  /// Returns row as Future [List] of [Cell].
  Future<List<Cell>> rowByKey(
    String key, {
    int column = 2,
    int length = -1,
  }) async {
    final row = await _ws.values.rowIndexOf(key, add: false);
    if (row < 1) return null;
    return this.row(row, column: column, length: length);
  }

  /// Fetches last column.
  ///
  /// [row] - optional (defaults to 1), index of a row that requested column
  /// starts from (cells before [row] will be skipped)
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all cells starting from [row] will be returned
  ///
  /// Returns last column as Future [List] of [Cell].
  Future<List<Cell>> lastColumn({
    int row = 1,
    int length = -1,
  }) async {
    final column = (await _ws.values.row(1)).length;
    return this.column(column, row: row, length: length);
  }

  /// Fetches last row.
  ///
  /// [column] - optional (defaults to 1), index of a column that requested row
  /// starts from (cells before [column] will be skipped)
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all cells starting from [column] will be returned
  ///
  /// Returns last row as Future [List] of [Cell].
  Future<List<Cell>> lastRow({
    int column = 1,
    int length = -1,
  }) async {
    final row = (await _ws.values.column(1)).length;
    return this.row(row, column: column, length: length);
  }

  /// Fetches all columns.
  ///
  /// [column] - optional (defaults to 1), index of a first returned column
  /// (columns before [column] will be skipped)
  /// columns start at index 1 (column A)
  ///
  /// [row] - optional (defaults to 1), index of a row that columns start from
  /// (cells before [row] will be skipped)
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested columns
  /// if length is `-1`, all cells starting from [row] will be returned
  ///
  /// Returns all columns as Future [List] of [List].
  Future<List<List<Cell>>> allColumns({
    int column = 1,
    int row = 1,
    int length = -1,
  }) async {
    final list = <List<Cell>>[];
    int colIndex = column;
    (await _ws.values.allColumns(column: column, row: row, length: length))
        .forEach((sublist) {
      int rowIndex = row;
      list.add(List.unmodifiable(sublist.map((value) {
        return Cell._(
          _ws,
          rowIndex++,
          colIndex,
          value,
        );
      })));
      colIndex++;
    });
    return List.unmodifiable(list);
  }

  /// Fetches all rows.
  ///
  /// [row] - optional (defaults to 1), index of a first returned row
  /// (rows before [row] will be skipped)
  /// rows start at index 1
  ///
  /// [column] - optional (defaults to 1), index of a column that rows start
  /// from (cells before [column] will be skipped)
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested rows
  /// if length is `-1`, all cells starting from [column] will be returned
  ///
  /// Returns all rows as Future [List] of [List].
  Future<List<List<Cell>>> allRows({
    int row = 1,
    int column = 1,
    int length = -1,
  }) async {
    final list = <List<Cell>>[];
    int rowIndex = row;
    (await _ws.values.allRows(row: row, column: column, length: length))
        .forEach((sublist) {
      int colIndex = column;
      list.add(List.unmodifiable(sublist.map((value) {
        return Cell._(
          _ws,
          rowIndex,
          colIndex++,
          value,
        );
      })));
      rowIndex++;
    });
    return List.unmodifiable(list);
  }

  /// Find cells by value.
  ///
  /// [value] - value to look for
  ///
  /// Returns cells as Future [List].
  Future<List<Cell>> findByValue(String value) async {
    final cells = <Cell>[];
    final rows = await _ws.values.allRows();
    int rowIndex = 1;
    for (List<String> row in rows) {
      int colIndex = 1;
      for (String val in row) {
        if (val == value) {
          cells.add(Cell._(_ws, rowIndex, colIndex, val));
        }
        colIndex++;
      }
      rowIndex++;
    }
    return cells;
  }

  /// Fetches cell.
  ///
  /// [column] - column index of a requested cell
  /// columns start at index 1 (column A)
  ///
  /// [row] - row index of a requested cell
  /// rows start at index 1
  ///
  /// Returns Future [Cell].
  Future<Cell> cell({
    @required int row,
    @required int column,
  }) async {
    final value = await _ws.values.value(column: column, row: row);
    return Cell._(_ws, row, column, value);
  }

  /// Fetches cell by names of its column and row.
  ///
  /// [rowKey] - name of a row with requested cell
  /// The column A considered to be row names
  ///
  /// [columnKey] - name of a column with requested cell
  /// The first row considered to be column names
  ///
  /// Returns Future [Cell].
  ///
  /// Returns `null` if either [rowKey] or [columnKey] not found.
  Future<Cell> cellByKeys({
    @required String rowKey,
    @required String columnKey,
  }) async {
    final column = await _ws.values.columnIndexOf(columnKey, add: false);
    if (column < 1) return null;
    final row = await _ws.values.rowIndexOf(rowKey, add: false);
    if (row < 1) return null;
    final value = await _ws.values.value(column: column, row: row);
    return Cell._(_ws, row, column, value);
  }

  /// Updates cells with values of [cells].
  ///
  /// [cells] - cells with values to insert (not null nor empty)
  ///
  /// Returns Future `true` in case of success.
  Future<bool> insert(List<Cell> cells) async {
    checkV(cells);
    final tuple = _cellsRangeTuple(cells);
    return _ws._update(
      values: cells.map((cell) => cell.value).toList(),
      range: tuple.first,
      majorDimension: tuple.second,
    );
  }

  Tuple<String, String> _cellsRangeTuple(List<Cell> cells) {
    final range =
        "'${cells.first.worksheetTitle}'!${cells.first.label}:${cells.last.label}";
    final dimen = cells.first.rowIndex == cells.last.rowIndex
        ? DIMEN_ROWS
        : DIMEN_COLUMNS;
    return Tuple(range, dimen);
  }
}

class _CellsMapper {
  final _WorksheetAsCells _cells;

  _CellsMapper(this._cells);

  /// Fetches specified column, maps it to other column and returns map.
  ///
  /// [column] - index of a requested column (values of returned map)
  /// columns start at index 1 (column A)
  ///
  /// [row] - optional (defaults to 1), index of a row that requested column
  /// starts from (cells before [row] will be skipped)
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all cells starting from [row] will be returned
  ///
  /// [mapTo] - optional (defaults to 1), index of a column to map cells to
  /// (keys of returned map)
  /// columns start at index 1 (column A)
  ///
  /// Returns column as Future [Map] of [String] to [Cell].
  Future<Map<String, Cell>> column(
    int column, {
    int row = 1,
    int mapTo = 1,
    int length = -1,
  }) async {
    checkM(column, mapTo);
    final values = await _cells.column(column, row: row, length: length);
    final keys =
        await _cells._ws.values.column(mapTo, row: row, length: length);
    final map = <String, Cell>{};
    mapKeysToValues(keys, values, map, null,
        (index) => Cell._(_cells._ws, row + index, column, ''));
    return Map.unmodifiable(map);
  }

  /// Fetches specified row, maps it to other row and returns map.
  ///
  /// [row] - index of a requested row (values of returned map)
  /// rows start at index 1
  ///
  /// [column] - optional (defaults to 1), index of a column that requested row
  /// starts from (cells before [column] will be skipped)
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all cells starting from [column] will be returned
  ///
  /// [mapTo] - optional (defaults to 1), index of a row to map cells to
  /// (keys of returned map)
  /// rows start at index 1
  ///
  /// Returns row as Future [Map] of [String] to [Cell].
  Future<Map<String, Cell>> row(
    int row, {
    int column = 1,
    int mapTo = 1,
    int length = -1,
  }) async {
    checkM(row, mapTo);
    final values = await _cells.row(row, column: column, length: length);
    final keys =
        await _cells._ws.values.row(mapTo, column: column, length: length);
    final map = <String, Cell>{};
    mapKeysToValues(keys, values, map, null,
        (index) => Cell._(_cells._ws, row, column + index, ''));
    return Map.unmodifiable(map);
  }

  /// Fetches column by its name, maps it to other column and returns map.
  ///
  /// The first row considered to be column names
  ///
  /// [key] - name of a requested column (values of returned map)
  ///
  /// [row] - optional (defaults to 1), index of a row that requested column
  /// starts from (cells before [row] will be skipped)
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all cells starting from [row] will be returned
  ///
  /// [mapTo] - optional, name of a column to map cells to (keys of returned
  /// map), if [mapTo] is `null` then cells will be mapped to column A
  /// columns start at index 1 (column A)
  ///
  /// Returns column as Future [Map] of [String] to [Cell].
  Future<Map<String, Cell>> columnByKey(
    String key, {
    int row = 2,
    String mapTo,
    int length = -1,
  }) async {
    checkM(key, mapTo);
    final column = await _cells._ws.values.columnIndexOf(key, add: false);
    if (column < 1) return null;
    final mapToIndex = isNullOrEmpty(mapTo)
        ? 1
        : await _cells._ws.values.columnIndexOf(mapTo, add: false);
    return this.column(column, row: row, length: length, mapTo: mapToIndex);
  }

  /// Fetches row by its name, maps it to other row, and returns map.
  ///
  /// The column A considered to be row names
  ///
  /// [key] - name of a requested row (values of returned map)
  ///
  /// [column] - optional (defaults to 1), index of a column that requested row
  /// starts from (cells before [column] will be skipped)
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all cells starting from [column] will be returned
  ///
  /// [mapTo] - optional, name of a row to map cells to (keys of returned
  /// map), if [mapTo] is `null` then cells will be mapped to first row
  /// rows start at index 1
  ///
  /// Returns row as Future [Map] of [String] to [Cell].
  Future<Map<String, Cell>> rowByKey(
    String key, {
    int column = 2,
    String mapTo,
    int length = -1,
  }) async {
    checkM(key, mapTo);
    final row = await _cells._ws.values.rowIndexOf(key, add: false);
    if (row < 1) return null;
    final mapToIndex = isNullOrEmpty(mapTo)
        ? 1
        : await _cells._ws.values.rowIndexOf(mapTo, add: false);
    return this.row(row, column: column, length: length, mapTo: mapToIndex);
  }

  /// Updates cells with values of [map].
  ///
  /// [map] - map containing cells with values to insert (not null nor empty)
  ///
  /// Returns Future `true` in case of success.
  Future<bool> insert(Map<String, Cell> map) async {
    return _cells.insert(map.values.toList()..sort());
  }
}
