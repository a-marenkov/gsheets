import 'dart:convert';

import 'package:googleapis/sheets/v4.dart' as v4;
import 'package:googleapis_auth/auth.dart';
import 'package:googleapis_auth/auth_io.dart';
import 'package:http/http.dart';
import 'package:meta/meta.dart';

import 'utils.dart';

const _sheetsEndpoint = 'https://sheets.googleapis.com/v4/spreadsheets/';
const _filesEndpoint = 'https://www.googleapis.com/drive/v2/files/';

/// [Exception] that throws gsheets.
///
/// [cause] - exception message.
///
/// GSheetsException is thrown:
/// - in case of invalid arguments;
/// - in case of google api returning an error.
class GSheetsException implements Exception {
  final String cause;

  GSheetsException(this.cause);

  @override
  String toString() => 'GSheetsException: $cause';
}

/// Manages googleapis auth and [Spreadsheet] fetching.
class GSheets {
  Future<AutoRefreshingAuthClient> _client;
  final ServiceAccountCredentials _credentials;
  final List<String> _scopes;

  /// Creates an instance of [GSheets].
  ///
  /// [credentialsJson] - must be provided, it can be either a [Map] or a
  /// JSON map encoded as a [String].
  ///
  /// [impersonatedUser] - optional, used to set the user to impersonate
  ///
  /// [scopes] - optional (defaults to `[SpreadsheetsScope, DriveScope]`).
  GSheets(
    credentialsJson, {
    String impersonatedUser,
    List<String> scopes = const [
      v4.SheetsApi.SpreadsheetsScope,
      v4.SheetsApi.DriveScope,
    ],
  })  : _credentials = ServiceAccountCredentials.fromJson(
          credentialsJson,
          impersonatedUser: impersonatedUser,
        ),
        _scopes = scopes {
    client; // initializes client
  }

  /// Returns Future [AutoRefreshingAuthClient] - autorefreshing,
  /// authenticated HTTP client.
  Future<AutoRefreshingAuthClient> get client {
    _client ??= clientViaServiceAccount(_credentials, _scopes);
    return _client;
  }

  /// Fetches and returns Future [Spreadsheet].
  ///
  /// Requires SheetsApi.SpreadsheetsScope.
  ///
  /// Throws Exception if [GSheets]'s scopes does not include SpreadsheetsScope.
  /// Throws GSheetsException if does not have permission.
  Future<Spreadsheet> spreadsheet(String spreadsheetId) async {
    final client = await this.client.catchError((_) {
      // retry once on error
      _client = null;
      return this.client;
    });
    final response = await client.get('$_sheetsEndpoint$spreadsheetId');
    checkResponse(response);
    final sheets = (jsonDecode(response.body)['sheets'] as List)
        .map((sheetJson) => Worksheet._fromSheetJson(
              sheetJson,
              client,
              spreadsheetId,
            ))
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
  ) async {
    final response = await client.post(
      '$_sheetsEndpoint$spreadsheetId:batchUpdate',
      body: jsonEncode({'requests': requests}),
    );
    checkResponse(response);
    return response;
  }
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
    final response = await _client.get('$_sheetsEndpoint$id');
    if (response.statusCode == 200) {
      final sheets = (jsonDecode(response.body)['sheets'] as List)
          .map((sheetJson) => Worksheet._fromSheetJson(
                sheetJson,
                _client,
                id,
              ))
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
  Worksheet worksheetByTitle(String title) {
    return sheets.firstWhere(
      (sheet) => sheet._title == title,
      orElse: () => null,
    );
  }

  /// Returns [Worksheet] with [id].
  ///
  /// Returns `null` if [Worksheet] with [id] not found.
  Worksheet worksheetById(int id) {
    return sheets.firstWhere(
      (sheet) => sheet.id == id,
      orElse: () => null,
    );
  }

  /// Returns [Worksheet] with [index].
  ///
  /// Returns `null` if [Worksheet] with [index] not found.
  Worksheet worksheetByIndex(int index) {
    return sheets.firstWhere(
      (sheet) => sheet.index == index,
      orElse: () => null,
    );
  }

  /// Adds new [Worksheet] with specified [title], [rows] and [columns].
  ///
  /// [title] - title of a new [Worksheet].
  /// [rows] - optional (defaults to 1000), row count of a new [Worksheet].
  /// [columns] - optional (defaults to 26), column count of a new [Worksheet].
  ///
  /// Returns Future of created [Worksheet].
  ///
  /// Throws [GSheetsException] if sheet with [title] already exists, or
  /// if [rows] or [columns] value is invalid.
  Future<Worksheet> addWorksheet(
    String title, {
    int rows = 1000,
    int columns = 26,
  }) async {
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
    final addSheetJson = (jsonDecode(response.body)['replies'] as List)?.first;
    if (addSheetJson != null) {
      final ws = Worksheet._fromSheetJson(
        addSheetJson['addSheet'],
        _client,
        id,
      );
      sheets.forEach((sheet) => sheet._incrementIndex(ws.index - 1));
      sheets.add(ws);
    }
    return worksheetByTitle(title);
  }

  /// Copies [ws] with specified [title] and [index].
  ///
  /// Returns Future of created [Worksheet].
  ///
  /// Throws [GSheetsException] if sheet with [title] already exists,
  /// or [index] is invalid.
  Future<Worksheet> copyWorksheet(
    Worksheet ws,
    String title, {
    int index,
  }) async {
    except((index ?? 0) < 0, 'invalid index ($index)');
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
    final duplicateSheetJson =
        (jsonDecode(response.body)['replies'] as List)?.first;
    if (duplicateSheetJson != null) {
      final ws = Worksheet._fromSheetJson(
        duplicateSheetJson['duplicateSheet'],
        _client,
        id,
      );
      sheets.forEach((sheet) => sheet._incrementIndex(ws.index - 1));
      sheets.add(ws);
    }
    return worksheetByTitle(title);
  }

  /// Deletes [ws].
  ///
  /// Returns `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> deleteWorksheet(Worksheet ws) async {
    await GSheets._batchUpdate(_client, id, [
      {
        'deleteSheet': {'sheetId': ws.id}
      }
    ]);
    sheets.remove(ws);
    sheets.forEach((sheet) => sheet._decrementIndex(ws.index));
    return true;
  }

  /// Returns Future list of [Permission].
  ///
  /// Requires SheetsApi.DriveScope.
  ///
  /// Throws Exception if [GSheets]'s scopes does not include DriveScope.
  /// Throws GSheetsException if DriveScope is not configured.
  Future<List<Permission>> permissions() async {
    final response = await _client.get('$_filesEndpoint$id/permissions');
    checkResponse(response);
    return (jsonDecode(response.body)['items'] as List)
            ?.map((json) => Permission._fromJson(json))
            ?.toList() ??
        [];
  }

  /// Returns Future [Permission] by email.
  ///
  /// Requires SheetsApi.DriveScope.
  ///
  /// [email] email of requested permission.
  ///
  /// Returns `null` if [email] not found.
  ///
  /// Throws Exception if [GSheets]'s scopes does not include DriveScope.
  /// Throws GSheetsException if DriveScope is not configured.
  Future<Permission> permissionByEmail(String email) async {
    final response = await _client.get('$_filesEndpoint$id/permissions');
    checkResponse(response);
    return (jsonDecode(response.body)['items'] as List)
        ?.map((json) => Permission._fromJson(json))
        ?.firstWhere((it) => it.email == email, orElse: () => null);
  }

  /// Shares [Spreadsheet].
  ///
  /// Requires SheetsApi.DriveScope.
  ///
  /// [user] - the email address or domain name for the entity.
  /// [type] - the account type.
  /// [role] - the primary role for this user.
  /// [withLink] - whether the link is required for this permission.
  ///
  /// Returns Future of shared [Permission].
  ///
  /// Throws Exception if [GSheets]'s scopes does not include DriveScope.
  /// Throws GSheetsException if DriveScope is not configured.
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
  /// [id] - permission id to remove.
  ///
  /// Returns `true` in case of success.
  ///
  /// Throws Exception if [GSheets]'s scopes does not include DriveScope.
  /// Throws GSheetsException if DriveScope is not configured.
  Future<bool> revokePermissionById(String id) async {
    final response = await _client.delete(
      '$_filesEndpoint${this.id}/permissions/$id',
    );
    checkResponse(response);
    return response.statusCode == 204;
  }

  /// Revokes permission by [email].
  ///
  /// Prefer using `revokePermissionById` if your know the id.
  ///
  /// [email] - email to remove permission for.
  ///
  /// Returns `true` in case of success.
  ///
  /// Throws Exception if [GSheets]'s scopes does not include DriveScope.
  /// Throws GSheetsException if DriveScope is not configured.
  Future<bool> revokePermissionByEmail(String email) async {
    final permission = await permissionByEmail(email);
    if (permission == null) {
      return false;
    }
    return revokePermissionById(permission.id);
  }
}

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
  int _index;
  WorksheetAsCells _cells;
  WorksheetAsValues _values;
  int _rowCount;
  int _columnCount;

  /// Current count of available rows.
  int get rowCount => _rowCount;

  /// Current count of available columns.
  int get columnCount => _columnCount;

  /// Current index of of the sheet.
  int get index => _index;

  /// Current title of of the sheet.
  String get title => _title;

  /// Interactor for working with [Worksheet] cells as [String] values.
  WorksheetAsValues get values {
    _values ??= WorksheetAsValues._(this);
    return _values;
  }

  /// Interactor for working with [Worksheet] cells as [Cell] objects.
  WorksheetAsCells get cells {
    _cells ??= WorksheetAsCells._(this);
    return _cells;
  }

  Worksheet._(
    this._client,
    this.spreadsheetId,
    this.id,
    this._title,
    this._index, [
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

  void _incrementIndex(int index) {
    if (_index > index) ++_index;
  }

  void _decrementIndex(int index) {
    if (_index > index) --_index;
  }

  /// Updates title of this [Worksheet].
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> updateTitle(String title) async {
    if (_title == title || isNullOrEmpty(title)) {
      return false;
    }
    await GSheets._batchUpdate(_client, spreadsheetId, [
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
    _title = title;
    return true;
  }

  Future<bool> _deleteDimension(String dimen, int index, int count) async {
    except(count < 1, 'invalid count ($count)');
    await GSheets._batchUpdate(_client, spreadsheetId, [
      {
        'deleteDimension': {
          'range': {
            'sheetId': id,
            'dimension': dimen,
            'startIndex': index - 1,
            'endIndex': index - 1 + count,
          },
        }
      }
    ]);
    return true;
  }

  /// Deletes columns from [Worksheet].
  ///
  /// [column] - index of a column to delete,
  /// columns start at index 1 (column A)
  ///
  /// [count] - optional (defaults to 1), the number of columns to delete
  /// starting from [column]
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> deleteColumn(int column, {int count = 1}) async {
    checkC(column);
    final isDeleted = await _deleteDimension(DIMEN_COLUMNS, column, count);
    if (isDeleted) {
      _columnCount = _columnCount - count;
    }
    return isDeleted;
  }

  /// Deletes rows from [Worksheet].
  ///
  /// [row] - index of a row to delete,
  /// rows start at index 1
  ///
  /// [count] - optional (defaults to 1), the number of rows to delete
  /// starting from [row]
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> deleteRow(int row, {int count = 1}) async {
    checkR(row);
    final isDeleted = await _deleteDimension(DIMEN_ROWS, row, count);
    if (isDeleted) {
      _rowCount = _rowCount - count;
    }
    return isDeleted;
  }

  Future<bool> _insertDimension(
    String dimen,
    int index,
    int count,
    bool inheritFromBefore,
  ) async {
    except(count < 1, 'invalid count ($count)');
    await GSheets._batchUpdate(_client, spreadsheetId, [
      {
        'insertDimension': {
          'range': {
            'sheetId': id,
            'dimension': dimen,
            'startIndex': index - 1,
            'endIndex': index - 1 + count,
          },
          'inheritFromBefore': inheritFromBefore
        }
      }
    ]);
    return true;
  }

  /// Inserts new columns to [Worksheet].
  ///
  /// [column] - index of a column to insert,
  /// columns start at index 1 (column A)
  ///
  /// [count] - optional (defaults to 1), the number of columns to insert
  ///
  /// [inheritFromBefore] - optional (defaults to `false`), if true, tells the
  /// API to give the new columns the same properties as the prior
  /// column, otherwise the new columns acquire the properties of
  /// those that follow them,
  /// cannot be true if inserting a column with index 1
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insertColumn(
    int column, {
    int count = 1,
    bool inheritFromBefore = false,
  }) async {
    checkC(column);
    final isInserted = await _insertDimension(
      DIMEN_COLUMNS,
      column,
      count,
      inheritFromBefore,
    );
    if (isInserted) {
      _columnCount = _columnCount + count;
    }
    return isInserted;
  }

  /// Inserts new rows to [Worksheet].
  ///
  /// [row] - index of a row to insert,
  /// rows start at index 1
  ///
  /// [count] - optional (defaults to 1), the number of rows to insert
  ///
  /// [inheritFromBefore] - optional (defaults to `false`), if true, tells the
  /// API to give the new rows the same properties as the prior
  /// row, otherwise the new rows acquire the properties of
  /// those that follow them,
  /// cannot be true if inserting a row with index 1
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insertRow(
    int row, {
    int count = 1,
    bool inheritFromBefore = false,
  }) async {
    checkC(row);
    final isInserted = await _insertDimension(
      DIMEN_ROWS,
      row,
      count,
      inheritFromBefore,
    );
    if (isInserted) {
      _rowCount = _rowCount + count;
    }
    return isInserted;
  }

  Future<bool> _moveDimension(
    String dimen,
    int from,
    int to,
    int count,
  ) async {
    except(from == to, 'cannot move from $from to $to');
    except(from < 1, 'invalid from ($from)');
    except(to < 1, 'invalid to ($to)');
    except(count < 1, 'invalid count ($count)');
    // correct values for from > to
    final cFrom = from < to ? from : to;
    final cTo = from < to ? to : from + count - 1;
    final cCount = from < to ? count : from - to;
    await GSheets._batchUpdate(_client, spreadsheetId, [
      {
        'moveDimension': {
          'source': {
            'sheetId': id,
            'dimension': dimen,
            'startIndex': cFrom - 1,
            'endIndex': cFrom - 1 + cCount,
          },
          'destinationIndex': cTo
        }
      }
    ]);
    return true;
  }

  /// Moves columns.
  ///
  /// [from] - index of a first column to move,
  /// columns start at index 1 (column A)
  ///
  /// [count] - optional (defaults to 1), the number of columns to move
  ///
  /// [to] - new index of a last column moved
  /// must be in a range from [from] to [from] + [count]
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> moveColumn({
    @required int from,
    @required int to,
    int count = 1,
  }) async {
    return _moveDimension(DIMEN_COLUMNS, from, to, count);
  }

  /// Moves rows.
  ///
  /// [from] - index of a first row to move,
  /// rows start at index 1
  ///
  /// [count] - optional (defaults to 1), the number of rows to move
  ///
  /// [to] - new index of a last row moved
  /// must be in a range from [from] to [from] + [count]
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> moveRow({
    @required int from,
    @required int to,
    int count = 1,
  }) async {
    return _moveDimension(DIMEN_ROWS, from, to, count);
  }

  Future<List<String>> _get(String range, String dimension) async {
    final response = await _client.get(
      '$_sheetsEndpoint$spreadsheetId/values/$range?majorDimension=$dimension',
    );
    checkResponse(response);
    return ((jsonDecode(response.body)['values'] as List)?.first as List)
            ?.cast<String>() ??
        [];
  }

  Future<List<List<String>>> _getAll(String range, String dimension) async {
    final response = await _client.get(
      '$_sheetsEndpoint$spreadsheetId/values/$range?majorDimension=$dimension',
    );
    checkResponse(response);
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
          'range': range,
          'majorDimension': majorDimension,
          'values': [values],
        },
      ),
    );
    checkResponse(response);
    return response.statusCode == 200;
  }

  Future<bool> _clear(String range) async {
    final response = await _client.post(
      '$_sheetsEndpoint$spreadsheetId/values/$range:clear',
    );
    checkResponse(response);
    return response.statusCode == 200;
  }

  Future<String> _rowRange(int row, int column, [int length = -1]) async {
    final expand = _expand(row, column + length - 1);
    final label = getColumnLetter(column);
    final labelTo = length > 0 ? getColumnLetter(column + length - 1) : '';
    await expand;
    return "'$_title'!${label}${row}:${labelTo}${row}";
  }

  Future<String> _columnRange(int column, int row, [int length = -1]) async {
    final expand = _expand(row + length - 1, column);
    final label = getColumnLetter(column);
    final to = length > 0 ? '${row + length - 1}' : '';
    await expand;
    return "'$_title'!$label${row}:$label$to";
  }

  Future<String> _allColumnsRange(
    int column,
    int row, [
    int length = -1,
    int count = -1,
  ]) async {
    final expand = _expand(row + length - 1, column);
    final fromLabel = getColumnLetter(column);
    final toLabel = count > 0
        ? getColumnLetter(column + count - 1)
        : getColumnLetter(columnCount);
    final to = length > 0 ? row + length - 1 : rowCount;
    await expand;
    return "'$_title'!$fromLabel${row}:$toLabel$to";
  }

  Future<String> _allRowsRange(
    int row,
    int column, [
    int length = -1,
    int count = -1,
  ]) async {
    final expand = _expand(row, column + length - 1);
    final label = getColumnLetter(column);
    final toLabel = length > 0
        ? getColumnLetter(column + length - 1)
        : getColumnLetter(columnCount);
    final toRow = count > 0 ? row + count - 1 : rowCount;
    await expand;
    return "'$_title'!${label}${row}:$toLabel$toRow";
  }

  Future<bool> _expand(int rows, int cols) async {
    var changed = false;
    if (_rowCount < rows) {
      _rowCount = rows;
      changed = true;
    }
    if (_columnCount < cols) {
      _columnCount = cols;
      changed = true;
    }
    if (changed) {
      final response = await GSheets._batchUpdate(_client, spreadsheetId, [
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
      checkResponse(response);
    }
    return changed;
  }

  /// Expands [Worksheet] by adding new rows/columns to the end of the sheet.
  ///
  /// [rows] - optional (defaults to 0), number of rows to add
  ///
  /// [columns] - optional (defaults to 0), number of columns to add
  ///
  /// Returns Future `true` if any rows/columns were added.
  ///
  /// Throws [GSheetsException].
  Future<bool> add({int rows = 0, int columns = 0}) async {
    except(rows < 0, 'invalid rows ($rows)');
    except(columns < 0, 'invalid column ($columns)');
    return _expand(_rowCount + rows, _columnCount + columns);
  }

  /// Clears the whole [Worksheet].
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> clear() => _clear("'$title'");

  /// Clears specified column.
  ///
  /// [column] - index of a column to clear,
  /// columns start at index 1 (column A)
  ///
  /// [fromRow] - optional (defaults to 1), index of a row that the column
  /// will be cleared from (values before [fromRow] will remain uncleared),
  /// rows start at index 1
  ///
  /// [count] - number of columns to clear
  ///
  /// [length] - optional (defaults to -1), number of cells to clear in the
  /// column
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> clearColumn(
    int column, {
    int fromRow = 1,
    int length = -1,
    int count = 1,
  }) async {
    checkCR(column, fromRow);
    except(count < 1, 'invalid count ($count)');
    return _clear(await _allColumnsRange(column, fromRow, length, count));
  }

  /// Clears specified row.
  ///
  /// [row] - index of a row to clear,
  /// rows start at index 1
  ///
  /// [fromColumn] - optional (defaults to 1), index of a column that the row
  /// will be cleared from (values before [fromColumn] will remain uncleared),
  /// columns start at index 1 (column A)
  ///
  /// [count] - number of rows to clear
  ///
  /// [length] - optional (defaults to -1), number of cells to clear in the row
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> clearRow(
    int row, {
    int fromColumn = 1,
    int length = -1,
    int count = 1,
  }) async {
    checkCR(fromColumn, row);
    except(count < 1, 'invalid count ($count)');
    return _clear(await _allRowsRange(row, fromColumn, length, count));
  }
}

/// Interactor for working with [Worksheet] cells as [String] values.
class WorksheetAsValues {
  final Worksheet _ws;
  ValueMapper _map;

  WorksheetAsValues._(this._ws);

  /// Mapper for [Worksheet]'s values.
  ValueMapper get map {
    _map ??= ValueMapper._(this);
    return _map;
  }

  /// Fetches specified column.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [column] - index of a requested column,
  /// columns start at index 1 (column A)
  ///
  /// [fromRow] - optional (defaults to 1), index of a row that requested column
  /// starts from (values before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all values starting from [fromRow] will be returned
  ///
  /// Returns column as Future [List] of [String].
  ///
  /// Throws [GSheetsException].
  Future<List<String>> column(
    int column, {
    int fromRow = 1,
    int length = -1,
  }) async {
    checkCR(column, fromRow);
    final range = await _ws._columnRange(column, fromRow, length);
    return _ws._get(range, DIMEN_COLUMNS);
  }

  /// Fetches specified row.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [row] - index of a requested row,
  /// rows start at index 1
  ///
  /// [fromColumn] - optional (defaults to 1), index of a column that requested row
  /// starts from (values before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all values starting from [fromColumn] will be returned
  ///
  /// Returns row as Future [List] of [String].
  ///
  /// Throws [GSheetsException].
  Future<List<String>> row(
    int row, {
    int fromColumn = 1,
    int length = -1,
  }) async {
    checkCR(fromColumn, row);
    final range = await _ws._rowRange(row, fromColumn, length);
    return _ws._get(range, DIMEN_ROWS);
  }

  /// Fetches column by its name.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [key] - name of a requested column
  /// The first row considered to be column names
  ///
  /// [fromRow] - optional (defaults to 2), index of a row that requested column
  /// starts from (values before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all values starting from [fromRow] will be returned
  ///
  /// Returns column as Future [List] of [String].
  ///
  /// Throws [GSheetsException].
  Future<List<String>> columnByKey(
    String key, {
    int fromRow = 2,
    int length = -1,
  }) async {
    final column = await columnIndexOf(key, add: false);
    if (column < 1) return null;
    return this.column(column, fromRow: fromRow, length: length);
  }

  /// Fetches row by its name.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [key] - name of a requested row
  /// The column A considered to be row names
  ///
  /// [fromColumn] - optional (defaults to 2), index of a column that requested row
  /// starts from (values before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all values starting from [fromColumn] will be returned
  ///
  /// Returns row as Future [List] of [String].
  ///
  /// Throws [GSheetsException].
  Future<List<String>> rowByKey(
    String key, {
    int fromColumn = 2,
    int length = -1,
  }) async {
    final row = await rowIndexOf(key, add: false);
    if (row < 1) return null;
    return this.row(row, fromColumn: fromColumn, length: length);
  }

  /// Fetches last column.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [fromRow] - optional (defaults to 1), index of a row that requested column
  /// starts from (values before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all values starting from [fromRow] will be returned
  ///
  /// Returns last column as Future [List] of [String].
  /// Returns Future `null` if there are no columns.
  ///
  /// Throws [GSheetsException].
  Future<List<String>> lastColumn({
    int fromRow = 1,
    int length = -1,
  }) async {
    final column = maxLength(await allRows());
    if (column < 1) return null;
    return this.column(column, fromRow: fromRow, length: length);
  }

  /// Fetches last row.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [fromColumn] - optional (defaults to 1), index of a column that requested row
  /// starts from (values before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all values starting from [fromColumn] will be returned
  ///
  /// Returns last row as Future [List] of [String].
  /// Returns Future `null` if there are no rows.
  ///
  /// Throws [GSheetsException].
  Future<List<String>> lastRow({
    int fromColumn = 1,
    int length = -1,
  }) async {
    final row = maxLength(await allColumns());
    if (row < 1) return null;
    return this.row(row, fromColumn: fromColumn, length: length);
  }

  /// Fetches all columns.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [fromColumn] - optional (defaults to 1), index of a first returned column
  /// (columns before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [fromRow] - optional (defaults to 1), index of a row that columns start from
  /// (values before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested columns
  /// if length is `-1`, all values starting from [fromRow] will be returned
  ///
  /// Returns all columns as Future [List] of [List].
  ///
  /// Throws [GSheetsException].
  Future<List<List<String>>> allColumns({
    int fromColumn = 1,
    int fromRow = 1,
    int length = -1,
  }) async {
    checkCR(fromColumn, fromRow);
    final range = await _ws._allColumnsRange(fromColumn, fromRow, length);
    return _ws._getAll(range, DIMEN_COLUMNS);
  }

  /// Fetches all rows.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [fromRow] - optional (defaults to 1), index of a first returned row
  /// (rows before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [fromColumn] - optional (defaults to 1), index of a column that rows start
  /// from (values before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested rows
  /// if length is `-1`, all values starting from [fromColumn] will be returned
  ///
  /// Returns all rows as Future [List] of [List].
  ///
  /// Throws [GSheetsException].
  Future<List<List<String>>> allRows({
    int fromRow = 1,
    int fromColumn = 1,
    int length = -1,
  }) async {
    checkCR(fromColumn, fromRow);
    final range = await _ws._allRowsRange(fromRow, fromColumn, length);
    return _ws._getAll(range, DIMEN_ROWS);
  }

  /// Fetches cell's value.
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [column] - column index of a requested cell's value,
  /// columns start at index 1 (column A)
  ///
  /// [row] - row index of a requested cell's value,
  /// rows start at index 1
  ///
  /// Returns cell's value as Future [String].
  ///
  /// Throws [GSheetsException].
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
  ///
  /// Throws [GSheetsException].
  Future<String> valueByKeys({
    @required String rowKey,
    @required String columnKey,
  }) async {
    final column = columnIndexOf(columnKey, add: false);
    final row = rowIndexOf(rowKey, add: false);
    if (await column < 1 || await row < 1) return null;
    final range = await _ws._columnRange(await column, await row);
    return getOrEmpty(await _ws._get(range, DIMEN_COLUMNS));
  }

  /// Updates cell's value to [value].
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [row] - row index to insert [value] to,
  /// rows start at index 1
  ///
  /// [column] - column index to insert [value] to,
  /// columns start at index 1 (column A)
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insertValue(
    String value, {
    @required int column,
    @required int row,
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
  ///
  /// Throws [GSheetsException].
  Future<bool> insertValueByKeys(
    String value, {
    @required String columnKey,
    @required String rowKey,
  }) async {
    final column = columnIndexOf(columnKey, add: true);
    final row = rowIndexOf(rowKey, add: true);
    return insertValue(value, column: await column, row: await row);
  }

  /// Returns index of a column with [key] value in [inRow].
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [key] - value to look for
  ///
  /// [inRow] - optional (defaults to 1), row index in which [key] is looked for,
  /// rows start at index 1
  ///
  /// [add] - optional (defaults to `false`), whether the [key] should be added
  /// to [inRow] in case of absence
  ///
  /// Returns Future `-1` if not found and [add] is false.
  ///
  /// Throws [GSheetsException].
  Future<int> columnIndexOf(
    String key, {
    bool add = false,
    int inRow = 1,
  }) async {
    except(isNullOrEmpty(key), 'invalid key ($key)');
    final columnKeys = await row(inRow);
    var column = columnKeys.indexOf(key) + 1;
    if (column < 1) {
      column = -1;
      if (add) {
        column = columnKeys.length + 1;
        final isAdded = await _ws._update(
          values: [key],
          range: await _ws._columnRange(column, inRow),
          majorDimension: DIMEN_COLUMNS,
        );
        if (!isAdded) column = -1;
      }
    }
    return column;
  }

  /// Returns index of a row with [key] value in [inColumn].
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [key] - value to look for
  ///
  /// [inColumn] - optional (defaults to 1), column index in which [key] is
  /// looked for,
  /// columns start at index 1 (column A)
  ///
  /// [add] - optional (defaults to `false`), whether the [key] should be added
  /// to [inColumn] in case of absence
  ///
  /// Returns Future `-1` if [key] is not found and [add] is false.
  ///
  /// Throws [GSheetsException].
  Future<int> rowIndexOf(
    String key, {
    bool add = false,
    inColumn = 1,
  }) async {
    except(isNullOrEmpty(key), 'invalid key ($key)');
    final rowKeys = await column(inColumn);
    var row = rowKeys.indexOf(key) + 1;
    if (row < 1) {
      row = -1;
      if (add) {
        row = rowKeys.length + 1;
        final isAdded = await _ws._update(
          values: [key],
          range: await _ws._columnRange(inColumn, row),
          majorDimension: DIMEN_COLUMNS,
        );
        if (!isAdded) row = -1;
      }
    }
    return row;
  }

  /// Updates column values with [values].
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [column] - column index to insert [values] to,
  /// columns start at index 1 (column A)
  ///
  /// [fromRow] - optional (defaults to 1), row index for the first inserted value
  /// of [values],
  /// rows start at index 1
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insertColumn(
    int column,
    List<String> values, {
    int fromRow = 1,
  }) async {
    checkCR(column, fromRow);
    checkV(values);
    return _ws._update(
      values: values,
      range: await _ws._columnRange(column, fromRow, values.length),
      majorDimension: DIMEN_COLUMNS,
    );
  }

  /// Updates column by its name values with [values].
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [key] - name of a column to insert [values] to
  /// The first row considered to be column names
  ///
  /// [fromRow] - optional (defaults to 2), row index for the first inserted value
  /// of [values],
  /// rows start at index 1
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insertColumnByKey(
    String key,
    List<String> values, {
    int fromRow = 2,
  }) async {
    final column = await columnIndexOf(key, add: true);
    return insertColumn(column, values, fromRow: fromRow);
  }

  /// Appends column.
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// The column index that [values] will be inserted to is the column index
  /// after last cell in the first row
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [fromRow] - optional (defaults to 1), row index for the first inserted value
  /// of [values],
  /// rows start at index 1
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> appendColumn(
    List<String> values, {
    int fromRow = 1,
  }) async {
    final column = maxLength(await allRows()) + 1;
    return insertColumn(column, values, fromRow: fromRow);
  }

  /// Updates row values with [values].
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [row] - row index to insert [values] to,
  /// rows start at index 1
  ///
  /// [fromColumn] - optional (defaults to 1), column index for the first inserted
  /// value of [values],
  /// columns start at index 1 (column A)
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insertRow(
    int row,
    List<String> values, {
    int fromColumn = 1,
  }) async {
    checkCR(fromColumn, row);
    checkV(values);
    return _ws._update(
      values: values,
      range: await _ws._rowRange(row, fromColumn, values.length),
      majorDimension: DIMEN_ROWS,
    );
  }

  /// Updates row by its name values with [values].
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [key] - name of a row to insert [values] to
  /// The column A considered to be row names
  ///
  /// [fromColumn] - optional (defaults to 2), column index for the first inserted
  /// value of [values],
  /// columns start at index 1 (column A)
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insertRowByKey(
    String key,
    List<String> values, {
    int fromColumn = 2,
  }) async {
    var row = await rowIndexOf(key, add: true);
    return insertRow(row, values, fromColumn: fromColumn);
  }

  /// Appends row.
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// The row index that [values] will be inserted to is the row index after
  /// last cell in the column A
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [fromColumn] - optional (defaults to 1), column index for the first inserted
  /// value of [values],
  /// columns start at index 1 (column A)
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> appendRow(
    List<String> values, {
    int fromColumn = 1,
  }) async {
    final row = maxLength(await allColumns()) + 1;
    return insertRow(row, values, fromColumn: fromColumn);
  }
}

/// Mapper for [Worksheet]'s values.
class ValueMapper {
  final WorksheetAsValues _values;

  ValueMapper._(this._values);

  /// Fetches specified column, maps it to other column and returns map.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [column] - index of a requested column (values of returned map),
  /// columns start at index 1 (column A)
  ///
  /// [fromRow] - optional (defaults to 1), index of a row that requested column
  /// starts from (values before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all values starting from [fromRow] will be returned
  ///
  /// [mapTo] - optional (defaults to 1), index of a column to map values to
  /// (keys of returned map),
  /// columns start at index 1 (column A)
  ///
  /// Returns column as Future [Map] of [String] to [String].
  ///
  /// Throws [GSheetsException].
  Future<Map<String, String>> column(
    int column, {
    int fromRow = 1,
    int length = -1,
    int mapTo = 1,
  }) async {
    checkM(column, mapTo);
    final values = _values.column(
      column,
      fromRow: fromRow,
      length: length,
    );
    final keys = _values.column(
      mapTo,
      fromRow: fromRow,
      length: length,
    );
    final map = <String, String>{};
    mapKeysToValues(await keys, await values, map, '', null);
    return map;
  }

  /// Fetches specified row, maps it to other row and returns map.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [row] - index of a requested row (values of returned map),
  /// rows start at index 1
  ///
  /// [fromColumn] - optional (defaults to 1), index of a column that requested row
  /// starts from (values before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all values starting from [fromColumn] will be returned
  ///
  /// [mapTo] - optional (defaults to 1), index of a row to map values to
  /// (keys of returned map),
  /// rows start at index 1
  ///
  /// Returns row as Future [Map] of [String] to [String].
  ///
  /// Throws [GSheetsException].
  Future<Map<String, String>> row(
    int row, {
    int fromColumn = 1,
    int length = -1,
    int mapTo = 1,
  }) async {
    checkM(row, mapTo);
    final values = _values.row(
      row,
      fromColumn: fromColumn,
      length: length,
    );
    final keys = _values.row(
      mapTo,
      fromColumn: fromColumn,
      length: length,
    );
    final map = <String, String>{};
    mapKeysToValues(await keys, await values, map, '', null);
    return map;
  }

  /// Fetches column by its name, maps it to other column and returns map.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// The first row considered to be column names
  ///
  /// [key] - name of a requested column (values of returned map)
  ///
  /// [fromRow] - optional (defaults to 1), index of a row that requested column
  /// starts from (values before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all values starting from [fromRow] will be returned
  ///
  /// [mapTo] - optional, name of a column to map values to (keys of returned
  /// map), if [mapTo] is `null` then values will be mapped to column A,
  /// columns start at index 1 (column A)
  ///
  /// Returns column as Future [Map] of [String] to [String].
  ///
  /// Throws [GSheetsException].
  Future<Map<String, String>> columnByKey(
    String key, {
    int fromRow = 2,
    int length = -1,
    String mapTo,
  }) async {
    checkM(key, mapTo);
    final column = _values.columnIndexOf(key, add: false);
    final mapToIndex = isNullOrEmpty(mapTo)
        ? Future.value(1)
        : _values.columnIndexOf(mapTo, add: false);
    if (await column < 1 || await mapToIndex < 1) return null;
    return this.column(
      await column,
      fromRow: fromRow,
      length: length,
      mapTo: await mapToIndex,
    );
  }

  /// Fetches row by its name, maps it to other row, and returns map.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// The column A considered to be row names
  ///
  /// [key] - name of a requested row (values of returned map)
  ///
  /// [fromColumn] - optional (defaults to 1), index of a column that requested row
  /// starts from (values before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all values starting from [fromColumn] will be returned
  ///
  /// [mapTo] - optional, name of a row to map values to (keys of returned
  /// map), if [mapTo] is `null` then values will be mapped to first row,
  /// rows start at index 1
  ///
  /// Returns row as Future [Map] of [String] to [String].
  ///
  /// Throws [GSheetsException].
  Future<Map<String, String>> rowByKey(
    String key, {
    int fromColumn = 2,
    int length = -1,
    String mapTo,
  }) async {
    checkM(key, mapTo);
    final row = _values.rowIndexOf(key, add: false);
    final mapToIndex = isNullOrEmpty(mapTo)
        ? Future.value(1)
        : _values.rowIndexOf(mapTo, add: false);
    if (await row < 1 || await mapToIndex < 1) return null;
    return this.row(
      await row,
      fromColumn: fromColumn,
      length: length,
      mapTo: await mapToIndex,
    );
  }

  /// Fetches last column, maps it to other column and returns map.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [fromRow] - optional (defaults to 1), index of a row that requested column
  /// starts from (values before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all values starting from [fromRow] will be returned
  ///
  /// [mapTo] - optional (defaults to 1), index of a column to map values to
  /// (keys of returned map),
  /// columns start at index 1 (column A)
  ///
  /// Returns column as Future [Map] of [String] to [String].
  /// Returns Future `null` if there are less than 2 columns.
  ///
  /// Throws [GSheetsException].
  Future<Map<String, String>> lastColumn({
    int fromRow = 1,
    int length = -1,
    int mapTo = 1,
  }) async {
    final column = maxLength(await _values.allRows());
    if (column < 2) return null;
    except(mapTo > column, 'invalid mapTo ($mapTo) - out of table bounds');
    return this.column(column, fromRow: fromRow, length: length, mapTo: mapTo);
  }

  /// Fetches last row, maps it to other row and returns map.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [fromColumn] - optional (defaults to 1), index of a column that requested row
  /// starts from (values before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all values starting from [fromColumn] will be returned
  ///
  /// [mapTo] - optional (defaults to 1), index of a row to map values to
  /// (keys of returned map),
  /// rows start at index 1
  ///
  /// Returns row as Future [Map] of [String] to [String].
  /// Returns Future `null` if there are less than 2 rows.
  ///
  /// Throws [GSheetsException].
  Future<Map<String, String>> lastRow({
    int fromColumn = 1,
    int length = -1,
    int mapTo = 1,
  }) async {
    final row = maxLength(await _values.allColumns());
    if (row < 2) return null;
    except(mapTo > row, 'invalid mapTo ($mapTo) - out of table bounds');
    return this.row(row, fromColumn: fromColumn, length: length, mapTo: mapTo);
  }

  /// Updates column values with values from [map].
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [map] - map containing values to insert (not null nor empty)
  ///
  /// [column] - column index to insert values of [map] to,
  /// columns start at index 1 (column A)
  ///
  /// [fromRow] - optional (defaults to 1), row index for the first inserted value,
  /// rows start at index 1
  ///
  /// [mapTo] - optional (defaults to 1), index of a column to which
  /// keys of the [map] will be mapped to,
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
  ///
  /// Throws [GSheetsException].
  Future<bool> insertColumn(
    int column,
    Map<String, String> map, {
    int fromRow = 1,
    int mapTo = 1,
    bool appendMissing = false,
    bool overwrite = false,
  }) async {
    checkCR(column, fromRow);
    checkM(column, mapTo);
    checkV(map);
    final columnMap = Map.from(map);
    final rows = await _values.column(mapTo, fromRow: fromRow);

    final newColumn = <String>[];
    if (overwrite) {
      for (var row in rows) {
        newColumn.add(columnMap.remove(row) ?? '');
      }
    } else {
      for (var row in rows) {
        newColumn.add(columnMap.remove(row));
      }
    }

    if (appendMissing && columnMap.isNotEmpty) {
      final newKeys = <String>[];
      for (var entry in columnMap.entries) {
        newKeys.add(entry.key);
        newColumn.add(entry.value);
      }
      final newKeysInsertion = _values.insertColumn(
        mapTo,
        newKeys,
        fromRow: fromRow + rows.length,
      );
      final newColumnInsertion = _values.insertColumn(
        column,
        newColumn,
        fromRow: fromRow,
      );
      await newKeysInsertion;
      return newColumnInsertion;
    }
    return _values.insertColumn(
      column,
      newColumn,
      fromRow: fromRow,
    );
  }

  /// Updates column values with values from [map] by column names.
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// The first row considered to be column names
  ///
  /// [map] - map containing values to insert (not null nor empty)
  ///
  /// [key] - name of a requested column (values of returned map)
  ///
  /// [fromRow] - optional (defaults to 2), row index for the first inserted value,
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
  ///
  /// Throws [GSheetsException].
  Future<bool> insertColumnByKey(
    String key,
    Map<String, String> map, {
    int fromRow = 2,
    String mapTo,
    bool appendMissing = false,
    bool overwrite = false,
  }) async {
    checkM(key, mapTo);
    final column = _values.columnIndexOf(key, add: true);
    final mapToIndex = isNullOrEmpty(mapTo)
        ? Future.value(1)
        : _values.columnIndexOf(mapTo, add: false);
    return insertColumn(
      await column,
      map,
      fromRow: fromRow,
      mapTo: await mapToIndex,
      appendMissing: appendMissing,
      overwrite: overwrite,
    );
  }

  /// Appends column with values from [map].
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [map] - map containing values to insert (not null nor empty)
  ///
  /// [fromRow] - optional (defaults to 1), row index for the first inserted value,
  /// rows start at index 1
  ///
  /// [mapTo] - optional (defaults to 1), index of a column to which
  /// keys of the [map] will be mapped to,
  /// columns start at index 1 (column A)
  ///
  /// [appendMissing] - optional (defaults to `false`), whether keys of [map]
  /// (with its related values) that are not present in a [mapTo] column
  /// should be added
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> appendColumn(
    Map<String, String> map, {
    int fromRow = 1,
    int mapTo = 1,
    bool appendMissing = false,
  }) async {
    final column = maxLength(await _values.allRows()) + 1;
    return insertColumn(
      column,
      map,
      fromRow: fromRow,
      mapTo: mapTo,
      appendMissing: appendMissing,
      overwrite: false,
    );
  }

  /// Updates row values with values from [map].
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [map] - map containing values to insert (not null nor empty)
  ///
  /// [row] - row index to insert values of [map] to,
  /// rows start at index 1
  ///
  /// [fromColumn] - optional (defaults to 1), column index for the first inserted
  /// value,
  /// columns start at index 1 (column A)
  ///
  /// [mapTo] - optional (defaults to 1), index of a row to which
  /// keys of the [map] will be mapped to,
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
  ///
  /// Throws [GSheetsException].
  Future<bool> insertRow(
    int row,
    Map<String, String> map, {
    int fromColumn = 1,
    int mapTo = 1,
    bool appendMissing = false,
    bool overwrite = false,
  }) async {
    checkCR(fromColumn, row);
    checkM(row, mapTo);
    checkV(map);

    final rowMap = Map.from(map);
    final columns = await _values.row(mapTo, fromColumn: fromColumn);

    final newRow = <String>[];
    if (overwrite) {
      for (var column in columns) {
        newRow.add(rowMap.remove(column) ?? '');
      }
    } else {
      for (var column in columns) {
        newRow.add(rowMap.remove(column));
      }
    }

    if (appendMissing && rowMap.isNotEmpty) {
      final newKeys = <String>[];
      for (var entry in rowMap.entries) {
        newKeys.add(entry.key);
        newRow.add(entry.value);
      }
      final newKeysInsertion = _values.insertRow(
        mapTo,
        newKeys,
        fromColumn: fromColumn + columns.length,
      );
      final newRowInsertion = _values.insertRow(
        row,
        newRow,
        fromColumn: fromColumn,
      );
      await newKeysInsertion;
      return newRowInsertion;
    }

    return _values.insertRow(
      row,
      newRow,
      fromColumn: fromColumn,
    );
  }

  /// Updates row values with values from [map] by column names.
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// The column A considered to be row names
  ///
  /// [map] - map containing values to insert (not null nor empty)
  ///
  /// [key] - name of a requested row (values of returned map)
  ///
  /// [fromColumn] - optional (defaults to 2), column index for the first inserted
  /// value,
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
  ///
  /// Throws [GSheetsException].
  Future<bool> insertRowByKey(
    String key,
    Map<String, String> map, {
    int fromColumn = 2,
    String mapTo,
    bool appendMissing = false,
    bool overwrite = false,
  }) async {
    checkM(key, mapTo);
    final row = _values.rowIndexOf(key, add: true);
    final mapToIndex = isNullOrEmpty(mapTo)
        ? Future.value(1)
        : _values.rowIndexOf(mapTo, add: false);
    return insertRow(
      await row,
      map,
      fromColumn: fromColumn,
      mapTo: await mapToIndex,
      appendMissing: appendMissing,
      overwrite: overwrite,
    );
  }

  /// Appends row with values from [map].
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [map] - map containing values to insert (not null nor empty)
  ///
  /// [fromColumn] - optional (defaults to 1), column index for the first inserted
  /// value,
  /// columns start at index 1 (column A)
  ///
  /// [mapTo] - optional (defaults to 1), index of a row to which
  /// keys of the [map] will be mapped to,
  /// rows start at index 1
  ///
  /// [appendMissing] - optional (defaults to `false`), whether keys of [map]
  /// (with its related values) that are not present in a [mapTo] row
  /// should be added
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> appendRow(
    Map<String, String> map, {
    int fromColumn = 1,
    int mapTo = 1,
    bool appendMissing = false,
  }) async {
    final row = maxLength(await _values.allColumns()) + 1;
    return insertRow(
      row,
      map,
      fromColumn: fromColumn,
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
  String get label => '${getColumnLetter(columnIndex)}${rowIndex}';

  String get worksheetTitle => _ws._title;

  /// Updates value of a cell.
  ///
  /// Returns Future `true` in case of success
  ///
  /// Throws [GSheetsException].
  Future<bool> post(String value) async {
    final posted = await _ws._update(
      values: [value ?? ''],
      range: "'$worksheetTitle'!$label",
      majorDimension: DIMEN_COLUMNS,
    );
    if (posted) {
      this.value = value;
    }
    return posted;
  }

  /// Refreshes value of a cell.
  ///
  /// Returns Future `true` if value has been changed.
  ///
  /// Throws [GSheetsException].
  Future<bool> refresh() async {
    final before = value;
    final range = await _ws._columnRange(columnIndex, rowIndex);
    value = getOrEmpty(await _ws._get(range, DIMEN_COLUMNS));
    return before != value;
  }

  @override
  String toString() => "'$value' at $label";

  @override
  int compareTo(other) {
    return rowIndex + columnIndex - other.rowIndex - other.columnIndex;
  }
}

/// Interactor for working with [Worksheet] cells as [Cell] objects.
class WorksheetAsCells {
  final Worksheet _ws;
  CellsMapper _map;

  WorksheetAsCells._(this._ws);

  /// Mapper for [Worksheet]'s cells.
  CellsMapper get map {
    _map ??= CellsMapper._(this);
    return _map;
  }

  /// Fetches specified column.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [column] - index of a requested column,
  /// columns start at index 1 (column A)
  ///
  /// [fromRow] - optional (defaults to 1), index of a row that requested column
  /// starts from (cells before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all cells starting from [fromRow] will be returned
  ///
  /// Returns column as Future [List] of [Cell].
  ///
  /// Throws [GSheetsException].
  Future<List<Cell>> column(
    int column, {
    int fromRow = 1,
    int length = -1,
  }) async {
    var index = fromRow;
    return List.unmodifiable(
        (await _ws.values.column(column, fromRow: fromRow, length: length))
            .map((value) => Cell._(
                  _ws,
                  index++,
                  column,
                  value,
                )));
  }

  /// Fetches specified row.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [row] - index of a requested row,
  /// rows start at index 1
  ///
  /// [fromColumn] - optional (defaults to 1), index of a column that requested row
  /// starts from (cells before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all cells starting from [fromColumn] will be returned
  ///
  /// Returns row as Future [List] of [Cell].
  ///
  /// Throws [GSheetsException].
  Future<List<Cell>> row(
    int row, {
    int fromColumn = 1,
    int length = -1,
  }) async {
    var index = fromColumn;
    return List.unmodifiable((await _ws.values.row(
      row,
      fromColumn: fromColumn,
      length: length,
    ))
        .map((value) => Cell._(
              _ws,
              row,
              index++,
              value,
            )));
  }

  /// Fetches column by its name.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [key] - name of a requested column
  /// The first row considered to be column names
  ///
  /// [fromRow] - optional (defaults to 2), index of a row that requested column
  /// starts from (cells before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all cells starting from [fromRow] will be returned
  ///
  /// Returns column as Future [List] of [Cell].
  ///
  /// Throws [GSheetsException].
  Future<List<Cell>> columnByKey(
    String key, {
    int fromRow = 2,
    int length = -1,
  }) async {
    final column = await _ws.values.columnIndexOf(key, add: false);
    if (column < 1) return null;
    return this.column(column, fromRow: fromRow, length: length);
  }

  /// Fetches row by its name.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [key] - name of a requested row
  /// The column A considered to be row names
  ///
  /// [fromColumn] - optional (defaults to 2), index of a column that requested row
  /// starts from (cells before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all cells starting from [fromColumn] will be returned
  ///
  /// Returns row as Future [List] of [Cell].
  ///
  /// Throws [GSheetsException].
  Future<List<Cell>> rowByKey(
    String key, {
    int fromColumn = 2,
    int length = -1,
  }) async {
    final row = await _ws.values.rowIndexOf(key, add: false);
    if (row < 1) return null;
    return this.row(row, fromColumn: fromColumn, length: length);
  }

  /// Fetches last column.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [fromRow] - optional (defaults to 1), index of a row that requested column
  /// starts from (cells before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all cells starting from [fromRow] will be returned
  ///
  /// Returns last column as Future [List] of [Cell].
  /// Returns Future `null` if there are no columns.
  ///
  /// Throws [GSheetsException].
  Future<List<Cell>> lastColumn({
    int fromRow = 1,
    int length = -1,
  }) async {
    final column = maxLength(await _ws.values.allRows());
    if (column < 1) return null;
    return this.column(column, fromRow: fromRow, length: length);
  }

  /// Fetches last row.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [fromColumn] - optional (defaults to 1), index of a column that requested row
  /// starts from (cells before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all cells starting from [fromColumn] will be returned
  ///
  /// Returns last row as Future [List] of [Cell].
  /// Returns Future `null` if there are no rows.
  ///
  /// Throws [GSheetsException].
  Future<List<Cell>> lastRow({
    int fromColumn = 1,
    int length = -1,
  }) async {
    final row = maxLength(await _ws.values.allColumns());
    if (row < 1) return null;
    return this.row(row, fromColumn: fromColumn, length: length);
  }

  /// Fetches all columns.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [fromColumn] - optional (defaults to 1), index of a first returned column
  /// (columns before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [fromRow] - optional (defaults to 1), index of a row that columns start from
  /// (cells before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested columns
  /// if length is `-1`, all cells starting from [fromRow] will be returned
  ///
  /// Returns all columns as Future [List] of [List].
  ///
  /// Throws [GSheetsException].
  Future<List<List<Cell>>> allColumns({
    int fromColumn = 1,
    int fromRow = 1,
    int length = -1,
  }) async {
    final list = <List<Cell>>[];
    var colIndex = fromColumn;
    final lists = await _ws.values.allColumns(
      fromColumn: fromColumn,
      fromRow: fromRow,
      length: length,
    );
    lists.forEach((sublist) {
      var rowIndex = fromRow;
      list.add(List.unmodifiable(sublist.map((value) => Cell._(
            _ws,
            rowIndex++,
            colIndex,
            value,
          ))));
      colIndex++;
    });
    return List.unmodifiable(list);
  }

  /// Fetches all rows.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [fromRow] - optional (defaults to 1), index of a first returned row
  /// (rows before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [fromColumn] - optional (defaults to 1), index of a column that rows start
  /// from (cells before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested rows
  /// if length is `-1`, all cells starting from [fromColumn] will be returned
  ///
  /// Returns all rows as Future [List] of [List].
  ///
  /// Throws [GSheetsException].
  Future<List<List<Cell>>> allRows({
    int fromRow = 1,
    int fromColumn = 1,
    int length = -1,
  }) async {
    final list = <List<Cell>>[];
    var rowIndex = fromRow;
    final lists = await _ws.values.allRows(
      fromRow: fromRow,
      fromColumn: fromColumn,
      length: length,
    );
    lists.forEach((sublist) {
      var colIndex = fromColumn;
      list.add(List.unmodifiable(sublist.map((value) => Cell._(
            _ws,
            rowIndex,
            colIndex++,
            value,
          ))));
      rowIndex++;
    });
    return List.unmodifiable(list);
  }

  /// Find cells by value.
  ///
  /// [value] - value to look for
  ///
  /// Returns cells as Future [List].
  ///
  /// Throws [GSheetsException].
  Future<List<Cell>> findByValue(String value) async {
    final cells = <Cell>[];
    final rows = await _ws.values.allRows();
    var rowIndex = 1;
    for (var row in rows) {
      var colIndex = 1;
      for (var val in row) {
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
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [column] - column index of a requested cell,
  /// columns start at index 1 (column A)
  ///
  /// [row] - row index of a requested cell,
  /// rows start at index 1
  ///
  /// Returns Future [Cell].
  ///
  /// Throws [GSheetsException].
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
  ///
  /// Throws [GSheetsException].
  Future<Cell> cellByKeys({
    @required String rowKey,
    @required String columnKey,
  }) async {
    final column = _ws.values.columnIndexOf(columnKey, add: false);
    final row = _ws.values.rowIndexOf(rowKey, add: false);
    if (await column < 1 || await row < 1) return null;
    final range = await _ws._columnRange(await column, await row);
    final value = getOrEmpty(await _ws._get(range, DIMEN_COLUMNS));
    return Cell._(_ws, await row, await column, value);
  }

  /// Updates cells with values of [cells].
  ///
  /// [cells] - cells with values to insert (not null nor empty)
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
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

/// Mapper for [Worksheet]'s cells.
class CellsMapper {
  final WorksheetAsCells _cells;

  CellsMapper._(this._cells);

  /// Fetches specified column, maps it to other column and returns map.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [column] - index of a requested column (values of returned map),
  /// columns start at index 1 (column A)
  ///
  /// [fromRow] - optional (defaults to 1), index of a row that requested column
  /// starts from (cells before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all cells starting from [fromRow] will be returned
  ///
  /// [mapTo] - optional (defaults to 1), index of a column to map cells to
  /// (keys of returned map),
  /// columns start at index 1 (column A)
  ///
  /// Returns column as Future [Map] of [String] to [Cell].
  ///
  /// Throws [GSheetsException].
  Future<Map<String, Cell>> column(
    int column, {
    int fromRow = 1,
    int mapTo = 1,
    int length = -1,
  }) async {
    checkM(column, mapTo);
    final values = _cells.column(
      column,
      fromRow: fromRow,
      length: length,
    );
    final keys = _cells._ws.values.column(
      mapTo,
      fromRow: fromRow,
      length: length,
    );
    final map = <String, Cell>{};
    mapKeysToValues(await keys, await values, map, null,
        (index) => Cell._(_cells._ws, fromRow + index, column, ''));
    return Map.unmodifiable(map);
  }

  /// Fetches specified row, maps it to other row and returns map.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [row] - index of a requested row (values of returned map),
  /// rows start at index 1
  ///
  /// [fromColumn] - optional (defaults to 1), index of a column that requested row
  /// starts from (cells before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all cells starting from [fromColumn] will be returned
  ///
  /// [mapTo] - optional (defaults to 1), index of a row to map cells to
  /// (keys of returned map),
  /// rows start at index 1
  ///
  /// Returns row as Future [Map] of [String] to [Cell].
  ///
  /// Throws [GSheetsException].
  Future<Map<String, Cell>> row(
    int row, {
    int fromColumn = 1,
    int mapTo = 1,
    int length = -1,
  }) async {
    checkM(row, mapTo);
    final values = _cells.row(
      row,
      fromColumn: fromColumn,
      length: length,
    );
    final keys = _cells._ws.values.row(
      mapTo,
      fromColumn: fromColumn,
      length: length,
    );
    final map = <String, Cell>{};
    mapKeysToValues(await keys, await values, map, null,
        (index) => Cell._(_cells._ws, row, fromColumn + index, ''));
    return Map.unmodifiable(map);
  }

  /// Fetches column by its name, maps it to other column and returns map.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// The first row considered to be column names
  ///
  /// [key] - name of a requested column (values of returned map)
  ///
  /// [fromRow] - optional (defaults to 1), index of a row that requested column
  /// starts from (cells before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all cells starting from [fromRow] will be returned
  ///
  /// [mapTo] - optional, name of a column to map cells to (keys of returned
  /// map), if [mapTo] is `null` then cells will be mapped to column A,
  /// columns start at index 1 (column A)
  ///
  /// Returns column as Future [Map] of [String] to [Cell].
  ///
  /// Throws [GSheetsException].
  Future<Map<String, Cell>> columnByKey(
    String key, {
    int fromRow = 2,
    String mapTo,
    int length = -1,
  }) async {
    checkM(key, mapTo);
    final column = _cells._ws.values.columnIndexOf(key, add: false);
    final mapToIndex = isNullOrEmpty(mapTo)
        ? Future.value(1)
        : _cells._ws.values.columnIndexOf(mapTo, add: false);
    if (await column < 1 || await mapToIndex < 1) return null;
    return this.column(
      await column,
      fromRow: fromRow,
      length: length,
      mapTo: await mapToIndex,
    );
  }

  /// Fetches row by its name, maps it to other row, and returns map.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// The column A considered to be row names
  ///
  /// [key] - name of a requested row (values of returned map)
  ///
  /// [fromColumn] - optional (defaults to 1), index of a column that requested row
  /// starts from (cells before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all cells starting from [fromColumn] will be returned
  ///
  /// [mapTo] - optional, name of a row to map cells to (keys of returned
  /// map), if [mapTo] is `null` then cells will be mapped to first row,
  /// rows start at index 1
  ///
  /// Returns row as Future [Map] of [String] to [Cell].
  ///
  /// Throws [GSheetsException].
  Future<Map<String, Cell>> rowByKey(
    String key, {
    int fromColumn = 2,
    String mapTo,
    int length = -1,
  }) async {
    checkM(key, mapTo);
    final row = _cells._ws.values.rowIndexOf(key, add: false);
    final mapToIndex = isNullOrEmpty(mapTo)
        ? Future.value(1)
        : _cells._ws.values.rowIndexOf(mapTo, add: false);
    if (await row < 1 || await mapToIndex < 1) return null;
    return this.row(
      await row,
      fromColumn: fromColumn,
      length: length,
      mapTo: await mapToIndex,
    );
  }

  /// Fetches last column, maps it to other column and returns map.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [fromRow] - optional (defaults to 1), index of a row that requested column
  /// starts from (cells before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of a requested column
  /// if length is `-1`, all cells starting from [fromRow] will be returned
  ///
  /// [mapTo] - optional (defaults to 1), index of a column to map cells to
  /// (keys of returned map),
  /// columns start at index 1 (column A)
  ///
  /// Returns column as Future [Map] of [String] to [Cell].
  /// Returns Future `null` if there are less than 2 columns.
  ///
  /// Throws [GSheetsException].
  Future<Map<String, Cell>> lastColumn({
    int fromRow = 1,
    int mapTo = 1,
    int length = -1,
  }) async {
    final column = maxLength(await _cells._ws.values.allRows());
    if (column < 2) return null;
    checkM(column, mapTo);
    except(mapTo > column, 'invalid mapTo ($mapTo) - out of table bounds');
    final values = _cells.column(
      column,
      fromRow: fromRow,
      length: length,
    );
    final keys = _cells._ws.values.column(
      mapTo,
      fromRow: fromRow,
      length: length,
    );
    final map = <String, Cell>{};
    mapKeysToValues(await keys, await values, map, null,
        (index) => Cell._(_cells._ws, fromRow + index, column, ''));
    return Map.unmodifiable(map);
  }

  /// Fetches last row, maps it to other row and returns map.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [fromColumn] - optional (defaults to 1), index of a column that requested row
  /// starts from (cells before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a requested row
  /// if length is `-1`, all cells starting from [fromColumn] will be returned
  ///
  /// [mapTo] - optional (defaults to 1), index of a row to map cells to
  /// (keys of returned map),
  /// rows start at index 1
  ///
  /// Returns row as Future [Map] of [String] to [Cell].
  /// Returns Future `null` if there are less than 2 rows.
  ///
  /// Throws [GSheetsException].
  Future<Map<String, Cell>> lastRow({
    int fromColumn = 1,
    int mapTo = 1,
    int length = -1,
  }) async {
    final row = maxLength(await _cells._ws.values.allColumns());
    if (row < 2) return null;
    checkM(row, mapTo);
    except(mapTo > row, 'invalid mapTo ($mapTo) - out of table bounds');
    final values = _cells.row(
      row,
      fromColumn: fromColumn,
      length: length,
    );
    final keys = _cells._ws.values.row(
      mapTo,
      fromColumn: fromColumn,
      length: length,
    );
    final map = <String, Cell>{};
    mapKeysToValues(await keys, await values, map, null,
        (index) => Cell._(_cells._ws, row, fromColumn + index, ''));
    return Map.unmodifiable(map);
  }

  /// Updates cells with values of [map].
  ///
  /// [map] - map containing cells with values to insert (not null nor empty)
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insert(Map<String, Cell> map) async {
    return _cells.insert(map.values.toList()..sort());
  }
}
