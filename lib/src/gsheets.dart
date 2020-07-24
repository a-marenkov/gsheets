import 'dart:convert';

import 'package:googleapis/sheets/v4.dart' as v4;
import 'package:googleapis_auth/auth.dart';
import 'package:googleapis_auth/auth_io.dart';
import 'package:http/http.dart' as http;
import 'package:meta/meta.dart';
import 'dart:math';

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
  })  : _scopes = scopes,
        _credentials = ServiceAccountCredentials.fromJson(
          credentialsJson,
          impersonatedUser: impersonatedUser,
        ) {
    // pre-initializes client
    client;
  }

  /// Returns Future [AutoRefreshingAuthClient] - autorefreshing,
  /// authenticated HTTP client.
  Future<AutoRefreshingAuthClient> get client {
    _client ??= clientViaServiceAccount(_credentials, _scopes);
    return _client;
  }

  /// Creates a new [Spreadsheet], and returns it.
  ///
  /// Requires SheetsApi.SpreadsheetsScope.
  ///
  /// It's recommended to save [Spreadsheet] id once its created, it also can be
  /// shared with the user by email via method `share` of [Spreadsheet].
  ///
  /// [worksheetTitles] - optional (defaults to `['Sheet1']`), titles of the
  /// worksheets that will be created along with the spreadsheet
  ///
  /// [render] - determines how values should be rendered in the output.
  /// https://developers.google.com/sheets/api/reference/rest/v4/ValueRenderOption
  ///
  /// [input] - determines how input data should be interpreted.
  /// https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
  ///
  /// Throws Exception if [GSheets]'s scopes does not include SpreadsheetsScope.
  /// Throws GSheetsException if does not have permission.
  Future<Spreadsheet> createSpreadsheet(
    String title, {
    List<String> worksheetTitles = const <String>['Sheet1'],
    ValueRenderOption render = ValueRenderOption.unformatted_value,
    ValueInputOption input = ValueInputOption.user_entered,
  }) async {
    final client = await this.client.catchError((_) {
      // retry once on error
      _client = null;
      return this.client;
    });
    final worksheets = worksheetTitles
        ?.map((title) => {
              'properties': {
                'title': title,
                'sheetType': 'GRID',
                'gridProperties': {
                  'rowCount': defaultRowsCount,
                  'columnCount': defaultColumnCount,
                }
              },
            })
        ?.toList();
    final response = await client.post(_sheetsEndpoint,
        body: jsonEncode({
          'properties': {
            'title': title,
          },
          'sheets': worksheets,
        }));
    checkResponse(response);
    final renderOption = _parseRenderOption(render);
    final inputOption = _parseInputOption(input);
    final spreadsheetId = jsonDecode(response.body)['spreadsheetId'];
    final sheets = (jsonDecode(response.body)['sheets'] as List)
        .map((json) => Worksheet._fromJson(
              json,
              client,
              spreadsheetId,
              renderOption,
              inputOption,
            ))
        .toList();
    return Spreadsheet._(
      client,
      spreadsheetId,
      sheets,
      renderOption,
      inputOption,
    );
  }

  /// Fetches and returns Future [Spreadsheet].
  ///
  /// Requires SheetsApi.SpreadsheetsScope.
  ///
  /// [render] - determines how values should be rendered in the output.
  /// https://developers.google.com/sheets/api/reference/rest/v4/ValueRenderOption
  ///
  /// [input] - determines how input data should be interpreted.
  /// https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
  ///
  /// Throws Exception if [GSheets]'s scopes does not include SpreadsheetsScope.
  /// Throws GSheetsException if does not have permission.
  Future<Spreadsheet> spreadsheet(
    String spreadsheetId, {
    ValueRenderOption render = ValueRenderOption.unformatted_value,
    ValueInputOption input = ValueInputOption.user_entered,
  }) async {
    final client = await this.client.catchError((_) {
      // retry once on error
      _client = null;
      return this.client;
    });
    final response = await client.get('$_sheetsEndpoint$spreadsheetId');
    checkResponse(response);
    final renderOption = _parseRenderOption(render);
    final inputOption = _parseInputOption(input);
    final sheets = (jsonDecode(response.body)['sheets'] as List)
        .where(gridSheetsFilter)
        .map((json) => Worksheet._fromJson(
              json,
              client,
              spreadsheetId,
              renderOption,
              inputOption,
            ))
        .toList();
    return Spreadsheet._(
      client,
      spreadsheetId,
      sheets,
      renderOption,
      inputOption,
    );
  }

  static String _parseRenderOption(ValueRenderOption option) {
    switch (option) {
      case ValueRenderOption.formatted_value:
        return 'FORMATTED_VALUE';
      case ValueRenderOption.formula:
        return 'FORMULA';
      default:
        return 'UNFORMATTED_VALUE';
    }
  }

  static String _parseInputOption(ValueInputOption option) {
    switch (option) {
      case ValueInputOption.user_entered:
        return 'USER_ENTERED';
      default:
        return 'RAW';
    }
  }

  /// Applies one or more updates to the spreadsheet.
  /// [About batchUpdate](https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/batchUpdate)
  ///
  /// [client] - client that is used for request.
  ///
  /// [spreadsheetId] - the id of a spreadsheet to perform [requests] on.
  ///
  /// [requests] - list of valid requests to perform on the [Spreadsheet] with
  /// [spreadsheetId]
  /// Information about requests is available at [the official Google docs](https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#request)
  ///
  /// To create [requests] you can use official [googleapis library](https://pub.dev/packages/googleapis)
  ///
  /// Returns the [Response] of batchUpdate request.
  /// [About batchUpdate response](https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/response)
  ///
  /// Throws [GSheetsException]
  static Future<http.Response> batchUpdate(
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

enum ValueRenderOption { formatted_value, unformatted_value, formula }
enum ValueInputOption { user_entered, raw }

/// Representation of a [Spreadsheet], manages [Worksheet]s.
class Spreadsheet {
  final AutoRefreshingAuthClient _client;

  /// [Spreadsheet]'s id
  final String id;

  /// List of [Worksheet]s
  final List<Worksheet> sheets;

  /// Determines how values should be rendered in the output.
  /// https://developers.google.com/sheets/api/reference/rest/v4/ValueRenderOption
  final String renderOption;

  /// Determines how input data should be interpreted.
  /// https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
  final String inputOption;

  Spreadsheet._(
    this._client,
    this.id,
    this.sheets,
    this.renderOption,
    this.inputOption,
  );

  /// Applies one or more updates to the spreadsheet.
  /// [About batchUpdate](https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/batchUpdate)
  ///
  /// [requests] - list of valid requests to perform on the [Spreadsheet]
  /// Information about requests is available at [the official Google docs](https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#request)
  ///
  /// To create [requests] you can use official [googleapis library](https://pub.dev/packages/googleapis)
  ///
  /// Returns the [Response] of batchUpdate request.
  /// [About batchUpdate response](https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/response)
  ///
  /// Throws [GSheetsException]
  Future<http.Response> batchUpdate(List<Map<String, dynamic>> requests) =>
      GSheets.batchUpdate(_client, id, requests);

  /// Refreshes [Spreadsheet].
  ///
  /// Should be called if you believe, that spreadsheet has been changed
  /// by another user (such as added/deleted/renamed worksheets).
  ///
  /// Returns Future `true` in case of success.
  Future<bool> refresh() async {
    final response = await _client.get('$_sheetsEndpoint$id');
    if (response.statusCode == 200) {
      final newSheets = (jsonDecode(response.body)['sheets'] as List)
          .where(gridSheetsFilter)
          .map((json) => Worksheet._fromJson(
                json,
                _client,
                id,
                renderOption,
                inputOption,
              ))
          .toList();
      // removing deleted sheets
      final newIds = newSheets.map((s) => s.id).toSet();
      final oldIds = sheets.map((s) => s.id).toSet();
      final deleted = oldIds.difference(newIds);
      for (final id in deleted) {
        sheets.removeWhere((s) => s.id == id);
      }
      // adding and updating sheets
      for (final sheet in newSheets) {
        final changed = sheets.firstWhere(
          (s) => s.id == sheet.id,
          orElse: () => null,
        );
        if (changed == null) {
          // adding new sheet
          sheets.add(sheet);
        } else {
          // updating old sheet
          changed._title = sheet._title;
          changed._index = sheet._index;
          changed._rowCount = sheet._rowCount;
          changed._columnCount = sheet._columnCount;
        }
      }
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
    int rows = defaultRowsCount,
    int columns = defaultColumnCount,
  }) async {
    check('columns', columns);
    check('rows', rows);
    final response = await GSheets.batchUpdate(_client, id, [
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
    if (addSheetJson == null) return null;
    final ws = Worksheet._fromJson(
      addSheetJson['addSheet'],
      _client,
      id,
      renderOption,
      inputOption,
    );
    sheets.forEach((sheet) => sheet._incrementIndex(ws.index - 1));
    sheets.add(ws);
    return ws;
  }

  /// Copies [Worksheet] from another spreadsheet (the name of the copy will
  /// be "Copy of {title of copied sheet}").
  ///
  /// [spreadsheetId] - source spreadsheet id
  /// [sheetId] - id of the worksheet to copy
  ///
  /// Provided `credentialsJson` has to have permission to [spreadsheetId].
  ///
  /// Returns Future [Worksheet] in case of success.
  ///
  /// Throws [GSheetsException].
  Future<Worksheet> addFromSpreadsheet(
    String spreadsheetId,
    int sheetId,
  ) async {
    if (isNullOrEmpty(spreadsheetId) || spreadsheetId == id) {
      throw GSheetsException('invalid spreadsheetId ($spreadsheetId)');
    }
    final response = await _client.post(
      '$_sheetsEndpoint$spreadsheetId/sheets/$sheetId:copyTo',
      body: jsonEncode({'destinationSpreadsheetId': id}),
    );
    checkResponse(response);
    final json = {'properties': jsonDecode(response.body)};
    final ws = Worksheet._fromJson(
      json,
      _client,
      id,
      renderOption,
      inputOption,
    );
    sheets.forEach((sheet) => sheet._incrementIndex(ws.index - 1));
    sheets.add(ws);
    return ws;
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
    final response = await GSheets.batchUpdate(_client, id, [
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
    if (duplicateSheetJson == null) return null;
    final duplicate = Worksheet._fromJson(
      duplicateSheetJson['duplicateSheet'],
      _client,
      id,
      renderOption,
      inputOption,
    );
    sheets.forEach((sheet) => sheet._incrementIndex(duplicate.index - 1));
    sheets.add(duplicate);
    return duplicate;
  }

  /// Deletes [ws].
  ///
  /// Returns `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> deleteWorksheet(Worksheet ws) async {
    await GSheets.batchUpdate(_client, id, [
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
        <Permission>[];
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
        ?.firstWhere(
          (it) => it.email == email,
          orElse: () => null,
        );
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
  Future<Permission> share(
    String user, {
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
  /// The ID of this [Permission]. This is a unique identifier for the grantee.
  final String id;

  /// The "pretty" name of the value of the [Permission].
  final String name;

  /// The email address of the user or group to which this permission refers.
  final String email;

  /// The type of the grantee (user, group, domain, anyone).
  final String type;

  /// The role granted by this permission (owner, organizer, fileOrganizer,
  /// writer, commenter, reader).
  final String role;

  /// Whether the account associated with this permission has been deleted.
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
  final String renderOption;
  final String inputOption;
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
    this.renderOption,
    this.inputOption,
  ]);

  factory Worksheet._fromJson(
    Map<String, dynamic> sheetJson,
    AutoRefreshingAuthClient client,
    String sheetsId,
    String renderOption,
    String inputOption,
  ) {
    return Worksheet._(
      client,
      sheetsId,
      sheetJson['properties']['sheetId'],
      sheetJson['properties']['title'],
      sheetJson['properties']['index'],
      sheetJson['properties']['gridProperties']['rowCount'],
      sheetJson['properties']['gridProperties']['columnCount'],
      renderOption,
      inputOption,
    );
  }

  @override
  String toString() {
    return 'Worksheet{spreadsheetId: $spreadsheetId, id: $id, title: $_title, index: $_index, rowCount: $_rowCount, columnCount: $_columnCount}';
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
    await GSheets.batchUpdate(_client, spreadsheetId, [
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

  /// Copies this [Worksheet] to another spreadsheet (the name of the copy will
  /// be "Copy of [title]").
  ///
  /// [spreadsheetId] - destination spreadsheet id.
  ///
  /// Provided `credentialsJson` has to have permission to [spreadsheetId].
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> copyTo(String spreadsheetId) async {
    if (isNullOrEmpty(spreadsheetId) || spreadsheetId == this.spreadsheetId) {
      throw GSheetsException('invalid spreadsheetId ($spreadsheetId)');
    }
    final response = await _client.post(
      '$_sheetsEndpoint${this.spreadsheetId}/sheets/$id:copyTo',
      body: jsonEncode({'destinationSpreadsheetId': spreadsheetId}),
    );
    checkResponse(response);
    return response.statusCode == 200;
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

  Future<bool> _deleteDimension(String dimen, int index, int count) async {
    check('count', count);
    await GSheets.batchUpdate(_client, spreadsheetId, [
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
    check('column', column);
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
    check('row', row);
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
    check('count', count);
    await GSheets.batchUpdate(_client, spreadsheetId, [
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
    check('column', column);
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
    check('row', row);
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
    check('from', from);
    check('to', to);
    check('count', count);
    // correct values for from > to
    final cFrom = from < to ? from : to;
    final cTo = from < to ? to : from + count - 1;
    final cCount = from < to ? count : from - to;
    await GSheets.batchUpdate(_client, spreadsheetId, [
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

  Future<bool> _clear(String range) async {
    var encodedRange = Uri.encodeComponent(range);
    final response = await _client.post(
      '$_sheetsEndpoint$spreadsheetId/values/$encodedRange:clear',
    );
    checkResponse(response);
    return response.statusCode == 200;
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
    check('column', column);
    check('fromRow', fromRow);
    check('count', count);
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
    check('row', row);
    check('fromColumn', fromColumn);
    check('count', count);
    return _clear(await _allRowsRange(row, fromColumn, length, count));
  }

  void _incrementIndex(int index) {
    if (_index > index) ++_index;
  }

  void _decrementIndex(int index) {
    if (_index > index) --_index;
  }

  Future<List<String>> _get(String range, String dimension) async {
    var encodedRange = Uri.encodeComponent(range);
    final response = await _client.get(
      '$_sheetsEndpoint$spreadsheetId/values/$encodedRange?majorDimension=$dimension&valueRenderOption=$renderOption',
    );
    checkResponse(response);
    final list = (jsonDecode(response.body)['values'] as List)?.first as List;
    return list?.map(parseValue)?.toList() ?? <String>[];
  }

  Future<List<List<String>>> _getAll(String range, String dimension) async {
    var encodedRange = Uri.encodeComponent(range);
    final response = await _client.get(
      '$_sheetsEndpoint$spreadsheetId/values/$encodedRange?majorDimension=$dimension&valueRenderOption=$renderOption',
    );
    checkResponse(response);
    final values = jsonDecode(response.body)['values'] as List;
    if (values == null) return <List<String>>[];
    final list = <List<String>>[];
    for (var sublist in values) {
      list.add((sublist as List)?.map(parseValue)?.toList() ?? <String>[]);
    }
    return list;
  }

  Future<bool> _update({
    List<dynamic> values,
    String majorDimension,
    String range,
  }) async {
    checkNotNested(values);
    var encodedRange = Uri.encodeComponent(range);
    final response = await _client.put(
      '$_sheetsEndpoint$spreadsheetId/values/$encodedRange?valueInputOption=$inputOption',
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

  Future<bool> _updateAll({
    List<List<dynamic>> values,
    String majorDimension,
    String range,
  }) async {
    var encodedRange = Uri.encodeComponent(range);
    final response = await _client.put(
      '$_sheetsEndpoint$spreadsheetId/values/$encodedRange?valueInputOption=$inputOption',
      body: jsonEncode(
        {
          'range': range,
          'majorDimension': majorDimension,
          'values': values,
        },
      ),
    );
    checkResponse(response);
    return response.statusCode == 200;
  }

  Future<String> _columnRange(int column, int row, int length) async {
    final expand = _expand(row + length - 1, column);
    final label = getColumnLetter(column);
    final to = length > 0 ? '${row + length - 1}' : '';
    await expand;
    return "'$_title'!$label${row}:$label$to";
  }

  Future<String> _rowRange(int row, int column, int length) async {
    final expand = _expand(row, column + length - 1);
    final label = getColumnLetter(column);
    final labelTo = length > 0 ? getColumnLetter(column + length - 1) : '';
    await expand;
    return "'$_title'!${label}${row}:${labelTo}${row}";
  }

  Future<String> _allColumnsRange(
    int column,
    int row,
    int length,
    int count,
  ) async {
    final expand = _expand(
      max(row, row + length - 1),
      max(column, column + count - 1),
    );
    final fromLabel = getColumnLetter(column);
    final to = length > 0 ? row + length - 1 : gsheetsCellsLimit;
    await expand;
    final toLabel = count > 0 ? getColumnLetter(column + count - 1) : '';
    return "'$_title'!$fromLabel${row}:$toLabel$to";
  }

  Future<String> _allRowsRange(
    int row,
    int column,
    int length,
    int count,
  ) async {
    final expand = _expand(
      max(row, row + count - 1),
      max(column, column + length - 1),
    );
    final label = getColumnLetter(column);
    final toLabel = length > 0 ? getColumnLetter(column + length - 1) : '';
    await expand;
    final toRow = count > 0 ? row + count - 1 : gsheetsCellsLimit;
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
      final response = await GSheets.batchUpdate(_client, spreadsheetId, [
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
}

/// Interactor for working with [Worksheet] cells as [String] values.
class WorksheetAsValues {
  final Worksheet _ws;
  ValuesMapper _map;

  WorksheetAsValues._(this._ws);

  /// Mapper for [Worksheet]'s values.
  ValuesMapper get map {
    _map ??= ValuesMapper._(this);
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
    check('column', column);
    check('fromRow', fromRow);
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
    check('row', row);
    check('fromColumn', fromColumn);
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
  /// Returns Future `null` if there [key] is not found.
  ///
  /// Throws [GSheetsException].
  Future<List<String>> columnByKey(
    dynamic key, {
    int fromRow = 2,
    int length = -1,
  }) async {
    final cKey = parseKey(key);
    check('fromRow', fromRow);
    final columns = await allColumns();
    final columnIndex = whereFirst(columns, cKey);
    if (columnIndex < 0) return null;
    return extractSublist(
      columns[columnIndex],
      from: fromRow - 1,
      length: length,
    );
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
  /// Returns Future `null` if there [key] is not found.
  ///
  /// Throws [GSheetsException].
  Future<List<String>> rowByKey(
    dynamic key, {
    int fromColumn = 2,
    int length = -1,
  }) async {
    final rKey = parseKey(key);
    check('fromColumn', fromColumn);
    final rows = await allRows();
    final rowIndex = whereFirst(rows, rKey);
    if (rowIndex < 0) return null;
    return extractSublist(
      rows[rowIndex],
      from: fromColumn - 1,
      length: length,
    );
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
  /// [inRange] - optional (defaults to false), whether should be fetched last
  /// column in range (respects [fromRow] and [length]) or last column in table
  ///
  /// Returns last column as Future [List] of [String].
  /// Returns Future `null` if there are no columns.
  ///
  /// Throws [GSheetsException].
  Future<List<String>> lastColumn({
    int fromRow = 1,
    int length = -1,
    bool inRange = false,
  }) async {
    if (inRange) {
      final columns = await allColumns(
        fromRow: fromRow,
        length: length,
      );
      if (columns.isEmpty) return null;
      return columns.last;
    } else {
      check('fromRow', fromRow);
      final columns = await allColumns();
      if (columns.isEmpty) return null;
      return extractSublist(
        columns.last,
        from: fromRow - 1,
        length: length,
      );
    }
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
  /// [inRange] - optional (defaults to false), whether should be fetched last
  /// row in range (respects [fromColumn] and [length]) or last row in table
  ///
  /// Returns last row as Future [List] of [String].
  /// Returns Future `null` if there are no rows.
  ///
  /// Throws [GSheetsException].
  Future<List<String>> lastRow({
    int fromColumn = 1,
    int length = -1,
    bool inRange = false,
  }) async {
    if (inRange) {
      final rows = await allRows(
        fromColumn: fromColumn,
        length: length,
      );
      if (rows.isEmpty) return null;
      return rows.last;
    } else {
      check('fromColumn', fromColumn);
      final rows = await allRows();
      if (rows.isEmpty) return null;
      return extractSublist(
        rows.last,
        from: fromColumn - 1,
        length: length,
      );
    }
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
  /// [length] - optional (defaults to -1), the length of requested columns
  /// if length is `-1`, all values starting from [fromRow] will be returned
  ///
  /// [count] - optional (defaults to -1), the number of requested columns
  /// if count is `-1`, all columns starting from [fromColumn] will be returned
  ///
  /// Returns all columns as Future [List] of [List].
  ///
  /// Throws [GSheetsException].
  Future<List<List<String>>> allColumns({
    int fromColumn = 1,
    int fromRow = 1,
    int length = -1,
    int count = -1,
  }) async {
    check('fromColumn', fromColumn);
    check('fromRow', fromRow);
    final range = await _ws._allColumnsRange(
      fromColumn,
      fromRow,
      length,
      count,
    );
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
  /// [length] - optional (defaults to -1), the length of requested rows
  /// if length is `-1`, all values starting from [fromColumn] will be returned
  ///
  /// [count] - optional (defaults to -1), the number of requested rows
  /// if count is `-1`, all rows starting from [fromRow] will be returned
  ///
  /// Returns all rows as Future [List] of [List].
  ///
  /// Throws [GSheetsException].
  Future<List<List<String>>> allRows({
    int fromRow = 1,
    int fromColumn = 1,
    int length = -1,
    int count = -1,
  }) async {
    check('fromColumn', fromColumn);
    check('fromRow', fromRow);
    final range = await _ws._allRowsRange(
      fromRow,
      fromColumn,
      length,
      count,
    );
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
    check('column', column);
    check('row', row);
    final range = await _ws._columnRange(column, row, 1);
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
    @required dynamic rowKey,
    @required dynamic columnKey,
  }) async {
    final rKey = parseKey(rowKey, 'row');
    final cKey = parseKey(columnKey, 'column');
    final rows = await allRows();
    if (rows.isEmpty) return null;
    final columnIndex = rows.first.indexOf(cKey);
    if (columnIndex < 0) return null;
    final rowIndex = whereFirst(rows, rKey);
    if (rowIndex < 0) return null;
    return getOrEmpty(rows[rowIndex], columnIndex);
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
    dynamic value, {
    @required int column,
    @required int row,
  }) async {
    check('column', column);
    check('row', row);
    return _ws._update(
      values: [value ?? ''],
      range: await _ws._columnRange(column, row, 1),
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
  /// [eager] - optional (defaults to `true`), whether to add
  /// [rowKey]/[columnKey] if absent
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insertValueByKeys(
    dynamic value, {
    @required dynamic columnKey,
    @required dynamic rowKey,
    bool eager = true,
  }) async {
    final rKey = parseKey(rowKey, 'row');
    final cKey = parseKey(columnKey, 'column');
    final rows = await allRows();
    final row = _rowOf(rows, rKey, eager);
    final column = _columnOf(rows, cKey, eager);
    if (await row < 1) return false;
    if (await column < 1) return false;
    return _ws._update(
      values: [value ?? ''],
      range: await _ws._columnRange(await column, await row, 1),
      majorDimension: DIMEN_COLUMNS,
    );
  }

  Future<int> _rowOf(
    List<List<String>> rows,
    String key,
    bool eager,
  ) async {
    var row = whereFirst(rows, key) + 1;
    if (eager && row < 1) {
      row = max(rows.length + 1, 2);
      await _ws._update(
        values: [key],
        range: await _ws._columnRange(1, row, 1),
        majorDimension: DIMEN_COLUMNS,
      );
    }
    return row;
  }

  Future<int> _columnOf(
    List<List<String>> rows,
    String key,
    bool eager,
  ) async {
    var column = 0;
    if (rows.isNotEmpty) {
      column = rows.first.indexOf(key) + 1;
    }
    if (eager && column < 1) {
      column = maxLength(rows, 1) + 1;
      await _ws._update(
        values: [key],
        range: await _ws._columnRange(column, 1, 1),
        majorDimension: DIMEN_COLUMNS,
      );
    }
    return column;
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
    dynamic key, {
    bool add = false,
    int inRow = 1,
  }) async {
    final cKey = parseKey(key);
    final columnKeys = await row(inRow);
    var column = columnKeys.indexOf(cKey) + 1;
    if (column < 1) {
      column = -1;
      if (add) {
        await _ws._update(
          values: [cKey],
          range: await _ws._columnRange(columnKeys.length + 1, inRow, 1),
          majorDimension: DIMEN_COLUMNS,
        );
        column = columnKeys.length + 1;
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
    dynamic key, {
    bool add = false,
    inColumn = 1,
  }) async {
    final rKey = parseKey(key);
    final rowKeys = await column(inColumn);
    var row = rowKeys.indexOf(rKey) + 1;
    if (row < 1) {
      row = -1;
      if (add) {
        await _ws._update(
          values: [rKey],
          range: await _ws._columnRange(inColumn, rowKeys.length + 1, 1),
          majorDimension: DIMEN_COLUMNS,
        );
        row = rowKeys.length + 1;
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
    List<dynamic> values, {
    int fromRow = 1,
  }) async {
    check('column', column);
    check('fromRow', fromRow);
    checkValues(values);
    return _ws._update(
      values: values,
      range: await _ws._columnRange(column, fromRow, values.length),
      majorDimension: DIMEN_COLUMNS,
    );
  }

  /// Updates columns with [values].
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [column] - column index to insert [values] to,
  /// columns start at index 1 (column A)
  ///
  /// [fromRow] - optional (defaults to 1), row index for the first inserted
  /// value of [values], rows start at index 1
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insertColumns(
    int column,
    List<List<dynamic>> values, {
    int fromRow = 1,
  }) async {
    check('column', column);
    check('fromRow', fromRow);
    checkValues(values);
    return _ws._updateAll(
      values: values,
      range: await _ws._allColumnsRange(column, fromRow, -1, values.length),
      majorDimension: DIMEN_COLUMNS,
    );
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
    List<dynamic> values, {
    int fromColumn = 1,
  }) async {
    check('row', row);
    check('fromColumn', fromColumn);
    checkValues(values);
    return _ws._update(
      values: values,
      range: await _ws._rowRange(row, fromColumn, values.length),
      majorDimension: DIMEN_ROWS,
    );
  }

  /// Updates rows with [values].
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [row] - row index to insert [values] to, rows start at index 1
  ///
  /// [fromColumn] - optional (defaults to 1), column index for the first
  /// inserted value of [values], columns start at index 1 (column A)
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insertRows(
    int row,
    List<List<dynamic>> values, {
    int fromColumn = 1,
  }) async {
    check('row', row);
    check('fromColumn', fromColumn);
    checkValues(values);
    return _ws._updateAll(
      values: values,
      range: await _ws._allRowsRange(row, fromColumn, -1, values.length),
      majorDimension: DIMEN_ROWS,
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
  /// [eager] - optional (defaults to `true`), whether to add [key] if absent
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insertColumnByKey(
    dynamic key,
    List<dynamic> values, {
    int fromRow = 2,
    bool eager = true,
  }) async {
    final column = await columnIndexOf(key, add: eager);
    if (column < 1) return false;
    return insertColumn(column, values, fromRow: fromRow);
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
  /// [eager] - optional (defaults to `true`), whether to add [key] if absent
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insertRowByKey(
    dynamic key,
    List<dynamic> values, {
    int fromColumn = 2,
    bool eager = true,
  }) async {
    final row = await rowIndexOf(key, add: eager);
    if (row < 1) return false;
    return insertRow(row, values, fromColumn: fromColumn);
  }

  /// Appends column.
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [fromRow] - optional (defaults to 1), row index for the first inserted value
  /// of [values],
  /// rows start at index 1
  ///
  /// [inRange] - optional (defaults to false), whether [values] should be
  /// appended to last column in range (respects [fromRow]) or last column in table
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> appendColumn(
    List<dynamic> values, {
    int fromRow = 1,
    bool inRange = false,
  }) async {
    final columns = await allColumns(fromRow: inRange ? fromRow : 1);
    return insertColumn(columns.length + 1, values, fromRow: fromRow);
  }

  /// Appends row.
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [fromColumn] - optional (defaults to 1), column index for the first inserted
  /// value of [values],
  /// columns start at index 1 (column A)
  ///
  /// [inRange] - optional (defaults to false), whether [values] should be
  /// appended to last row in range (respects [fromColumn]) or last row in table
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> appendRow(
    List<dynamic> values, {
    int fromColumn = 1,
    bool inRange = false,
  }) async {
    final rows = await allRows(fromColumn: inRange ? fromColumn : 1);
    return insertRow(rows.length + 1, values, fromColumn: fromColumn);
  }

  /// Appends rows.
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [fromColumn] - optional (defaults to 1), column index for the first
  /// inserted value of [values], columns start at index 1 (column A)
  ///
  /// [inRange] - optional (defaults to false), whether [values] should be
  /// appended to last row in range (respects [fromColumn]) or last row in table
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> appendRows(
    List<List<dynamic>> values, {
    int fromColumn = 1,
    bool inRange = false,
  }) async {
    final rows = await allRows(fromColumn: inRange ? fromColumn : 1);
    return insertRows(rows.length + 1, values, fromColumn: fromColumn);
  }

  /// Appends columns.
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [values] - values to insert (not null nor empty)
  ///
  /// [fromRow] - optional (defaults to 1), row index for the first inserted
  /// value of [values], rows start at index 1
  ///
  /// [inRange] - optional (defaults to false), whether [values] should be
  /// appended to last column in range (respects [fromRow]) or last column in
  /// table
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> appendColumns(
    List<List<dynamic>> values, {
    int fromRow = 1,
    bool inRange = false,
  }) async {
    final columns = await allColumns(fromRow: inRange ? fromRow : 1);
    return insertColumns(columns.length + 1, values, fromRow: fromRow);
  }
}

/// Mapper for [Worksheet]'s values.
class ValuesMapper {
  final WorksheetAsValues _values;

  ValuesMapper._(this._values);

  Map<String, String> _wrap(List<String> keys, List<String> values) {
    return mapKeysToValues(keys, values, (_, val) => val ?? '');
  }

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
    check('column', column);
    check('mapTo', mapTo);
    checkMapTo(column, mapTo);
    final columns = await _values.allColumns(
      fromRow: fromRow,
      length: length,
    );
    final keys = get(columns, at: mapTo - 1, or: <String>[]);
    final values = get(columns, at: column - 1, or: <String>[]);
    return _wrap(keys, values);
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
    check('row', row);
    check('mapTo', mapTo);
    checkMapTo(row, mapTo);
    final rows = await _values.allRows(
      fromColumn: fromColumn,
      length: length,
    );
    final keys = get(rows, at: mapTo - 1, or: <String>[]);
    final values = get(rows, at: row - 1, or: <String>[]);
    return _wrap(keys, values);
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
    dynamic key, {
    int fromRow = 2,
    int length = -1,
    dynamic mapTo,
  }) async {
    final cKey = parseKey(key);
    final mKey = parseMapToKey(mapTo);
    check('fromRow', fromRow);
    checkMapTo(cKey, mKey);
    final columns = await _values.allColumns();
    if (columns.isEmpty) return null;
    final columnIndex = whereFirst(columns, cKey);
    if (columnIndex < 0) return null;
    final mapToIndex = isNullOrEmpty(mKey) ? 0 : whereFirst(columns, mKey);
    if (mapToIndex < 0) return null;
    checkMapTo(columnIndex + 1, mapToIndex + 1);
    final keys = extractSublist(
      columns[mapToIndex],
      from: fromRow - 1,
      length: length,
    );
    final values = extractSublist(
      columns[columnIndex],
      from: fromRow - 1,
      length: length,
    );
    return _wrap(keys, values);
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
    dynamic key, {
    int fromColumn = 2,
    int length = -1,
    dynamic mapTo,
  }) async {
    final rKey = parseKey(key);
    final mKey = parseMapToKey(mapTo);
    check('fromColumn', fromColumn);
    checkMapTo(rKey, mKey);
    final rows = await _values.allRows();
    if (rows.isEmpty) return null;
    final rowIndex = whereFirst(rows, rKey);
    if (rowIndex < 0) return null;
    final mapToIndex = isNullOrEmpty(mKey) ? 0 : whereFirst(rows, mKey);
    if (mapToIndex < 0) return null;
    checkMapTo(rowIndex + 1, mapToIndex + 1);
    final keys = extractSublist(
      rows[mapToIndex],
      from: fromColumn - 1,
      length: length,
    );
    final values = extractSublist(
      rows[rowIndex],
      from: fromColumn - 1,
      length: length,
    );
    return _wrap(keys, values);
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
  /// [inRange] - optional (defaults to false), whether should be fetched last
  /// column in range (respects [fromRow] and [length]) or last column in table
  ///
  /// Returns column as Future [Map] of [String] to [String].
  /// Returns Future `null` if there are less than 2 columns.
  ///
  /// Throws [GSheetsException].
  Future<Map<String, String>> lastColumn({
    int fromRow = 1,
    int length = -1,
    int mapTo = 1,
    bool inRange = false,
  }) async {
    check('mapTo', mapTo);
    if (inRange) {
      final columns = await _values.allColumns(
        fromRow: fromRow,
        length: length,
      );
      if (columns.length < 2) return null;
      checkMapTo(columns.length, mapTo);
      final column = columns.length;
      final keys = get(columns, at: mapTo - 1, or: <String>[]);
      final values = get(columns, at: column - 1, or: <String>[]);
      return _wrap(keys, values);
    } else {
      check('fromRow', fromRow);
      final columns = await _values.allColumns();
      if (columns.length < 2) return null;
      checkMapTo(columns.length, mapTo);
      final column = columns.length;
      final keys = extractSublist(
        get(columns, at: mapTo - 1, or: <String>[]),
        from: fromRow - 1,
        length: length,
      );
      final values = extractSublist(
        get(columns, at: column - 1, or: <String>[]),
        from: fromRow - 1,
        length: length,
      );
      return _wrap(keys, values);
    }
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
  /// [inRange] - optional (defaults to false), whether should be fetched last
  /// row in range (respects [fromColumn] and [length]) or last row in table
  ///
  /// Returns row as Future [Map] of [String] to [String].
  /// Returns Future `null` if there are less than 2 rows.
  ///
  /// Throws [GSheetsException].
  Future<Map<String, String>> lastRow({
    int fromColumn = 1,
    int length = -1,
    int mapTo = 1,
    bool inRange = false,
  }) async {
    check('mapTo', mapTo);
    if (inRange) {
      final rows = await _values.allRows(
        fromColumn: fromColumn,
        length: length,
      );
      if (rows.length < 2) return null;
      checkMapTo(rows.length, mapTo);
      final row = rows.length;
      final keys = get(rows, at: mapTo - 1, or: <String>[]);
      final values = get(rows, at: row - 1, or: <String>[]);
      return _wrap(keys, values);
    } else {
      check('fromColumn', fromColumn);
      final rows = await _values.allRows();
      if (rows.length < 2) return null;
      checkMapTo(rows.length, mapTo);
      final row = rows.length;
      final keys = extractSublist(
        get(rows, at: mapTo - 1, or: <String>[]),
        from: fromColumn - 1,
        length: length,
      );
      final values = extractSublist(
        get(rows, at: row - 1, or: <String>[]),
        from: fromColumn - 1,
        length: length,
      );
      return _wrap(keys, values);
    }
  }

  /// Fetches all columns, maps them to specific column and returns as list of
  /// maps.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [fromColumn] - optional (defaults to 1), index of a first returned column
  /// (columns before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [count] - optional (defaults to -1), the number of requested columns
  /// if count is `-1`, all columns starting from [fromColumn] will be returned
  ///
  /// [fromRow] - optional (defaults to 1), index of a row that requested
  /// columns start from (values before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [length] - optional (defaults to -1), the length of requested columns
  /// if length is `-1`, all values starting from [fromRow] will be returned
  ///
  /// [mapTo] - optional (defaults to 1), index of a column to map values to
  /// (keys of returned maps),
  /// columns start at index 1 (column A)
  ///
  /// Returns columns as Future `List<Map<String, String>>`.
  /// Returns Future `null` if there are less than 2 columns.
  ///
  /// Throws [GSheetsException].
  Future<List<Map<String, String>>> allColumns({
    int fromColumn = 1,
    int fromRow = 1,
    int length = -1,
    int count = -1,
    int mapTo = 1,
  }) async {
    check('fromColumn', fromColumn);
    check('mapTo', mapTo);
    final columns = await _values.allColumns(
      fromRow: fromRow,
      length: length,
    );
    if (columns.length < 2) return null;
    final maps = <Map<String, String>>[];
    final keys = get(columns, at: mapTo - 1, or: <String>[]);
    if (keys.isEmpty) return maps;
    final start = min(fromColumn - 1, columns.length);
    final end = count < 1 ? columns.length : min(start + count, columns.length);
    for (var i = start; i < end; i++) {
      if (i != mapTo - 1) {
        maps.add(_wrap(keys, columns[i]));
      }
    }
    return maps;
  }

  /// Fetches all rows, maps them to specific row and returns as list of
  /// maps.
  ///
  /// Expands current sheet's size if requested range is out of sheet's bounds.
  ///
  /// [fromRow] - optional (defaults to 1), index of a first returned row
  /// (rows before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [count] - optional (defaults to -1), the number of requested rows
  /// if count is `-1`, all rows starting from [fromRow] will be returned
  ///
  /// [fromColumn] - optional (defaults to 1), index of a column that requested
  /// rows start from (values before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of requested rows
  /// if length is `-1`, all values starting from [fromColumn] will be returned
  ///
  /// [mapTo] - optional (defaults to 1), index of a row to map values to
  /// (keys of returned maps),
  /// rows start at index 1
  ///
  /// Returns rows as Future `List<Map<String, String>>`.
  /// Returns Future `null` if there are less than 2 rows.
  ///
  /// Throws [GSheetsException].
  Future<List<Map<String, String>>> allRows({
    int fromRow = 1,
    int fromColumn = 1,
    int length = -1,
    int count = -1,
    int mapTo = 1,
  }) async {
    check('fromRow', fromRow);
    check('mapTo', mapTo);
    final rows = await _values.allRows(
      fromColumn: fromColumn,
      length: length,
    );
    if (rows.length < 2) return null;
    final maps = <Map<String, String>>[];
    final keys = get(rows, at: mapTo - 1, or: <String>[]);
    if (keys.isEmpty) return maps;
    final start = min(fromRow - 1, rows.length);
    final end = count < 1 ? rows.length : min(start + count, rows.length);
    for (var i = start; i < end; i++) {
      if (i != mapTo - 1) {
        maps.add(_wrap(keys, rows[i]));
      }
    }
    return maps;
  }

  Future<bool> _insertColumns(
    int column,
    List<Map<String, dynamic>> maps,
    List<String> allKeys,
    int fromRow,
    int mapTo,
    bool appendMissing,
    bool overwrite,
  ) async {
    final keys = extractSublist(allKeys, from: fromRow - 1);
    final newKeys = <String>{};
    final columns = <List>[];
    var newKeysInsertion = Future.value(true);
    if (appendMissing) {
      for (var map in maps) {
        newKeys.addAll(map.keys);
      }
      newKeys.removeAll(allKeys);
      if (newKeys.isNotEmpty) {
        newKeysInsertion = _values.insertColumn(
          mapTo,
          newKeys.toList(),
          fromRow: fromRow + keys.length,
        );
      }
    }
    if (overwrite) {
      for (var map in maps) {
        final column = [];
        columns.add(column);
        for (var row in keys) {
          column.add(map[row] ?? '');
        }
        for (var row in newKeys) {
          column.add(map[row] ?? '');
        }
      }
    } else {
      for (var map in maps) {
        final column = [];
        columns.add(column);
        for (var row in keys) {
          column.add(map[row]);
        }
        for (var row in newKeys) {
          column.add(map[row]);
        }
      }
    }
    await newKeysInsertion;
    return _values.insertColumns(
      column,
      columns,
      fromRow: fromRow,
    );
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
    Map<String, dynamic> map, {
    int fromRow = 1,
    int mapTo = 1,
    bool appendMissing = false,
    bool overwrite = false,
  }) async {
    checkMap(map);
    return insertColumns(
      column,
      [map],
      fromRow: fromRow,
      mapTo: mapTo,
      appendMissing: appendMissing,
      overwrite: overwrite,
    );
  }

  /// Updates columns with values from [maps].
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [maps] - list of maps containing values to insert (not null nor empty)
  ///
  /// [column] - column index to insert values of [maps] to,
  /// columns start at index 1 (column A)
  ///
  /// [fromRow] - optional (defaults to 1), row index for the first inserted value,
  /// rows start at index 1
  ///
  /// [mapTo] - optional (defaults to 1), index of a column to which
  /// keys of the [maps] will be mapped to,
  /// columns start at index 1 (column A)
  ///
  /// [appendMissing] - optional (defaults to `false`), whether keys of [maps]
  /// (with its related values) that are not present in a [mapTo] column
  /// should be added
  ///
  /// [overwrite] - optional (defaults to `false`), whether clear cells of
  /// [column] if [maps] does not contain value for them
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insertColumns(
    int column,
    List<Map<String, dynamic>> maps, {
    int fromRow = 1,
    int mapTo = 1,
    bool appendMissing = false,
    bool overwrite = false,
  }) async {
    check('column', column);
    checkMapTo(row, mapTo);
    checkMap(maps);
    final keys = await _values.column(mapTo);
    return _insertColumns(
      column,
      maps,
      keys,
      fromRow,
      mapTo,
      appendMissing,
      overwrite,
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
  /// [eager] - optional (defaults to `true`), whether to add [key] if absent
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insertColumnByKey(
    dynamic key,
    Map<String, dynamic> map, {
    int fromRow = 2,
    dynamic mapTo,
    bool appendMissing = false,
    bool overwrite = false,
    bool eager = true,
  }) async {
    final cKey = parseKey(key);
    final mKey = parseMapToKey(mapTo);
    check('fromRow', fromRow);
    checkMapTo(cKey, mapTo);
    checkMap(map);
    final columns = await _values.allColumns();
    final mapToIndex = isNullOrEmpty(mKey) ? 0 : whereFirst(columns, mKey);
    if (mapToIndex < 0) return false;
    var columnIndex = whereFirst(columns, cKey);
    if (columnIndex < 0) {
      if (!eager || columns.isEmpty || columns.first.isEmpty) return false;
      columnIndex = columns.length;
      await _values._ws._update(
        values: [cKey],
        range: await _values._ws._columnRange(columnIndex + 1, 1, 1),
        majorDimension: DIMEN_COLUMNS,
      );
    } else {
      checkMapTo(columnIndex + 1, mapToIndex + 1);
    }
    final keys = get(columns, at: mapToIndex, or: <String>[]);
    return _insertColumns(
      columnIndex + 1,
      [map],
      keys,
      fromRow,
      mapToIndex + 1,
      appendMissing,
      overwrite,
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
  /// [inRange] - optional (defaults to false), whether [map] values should be
  /// appended to last column in range (respects [fromRow]) or last column in table
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> appendColumn(
    Map<String, dynamic> map, {
    int fromRow = 1,
    int mapTo = 1,
    bool appendMissing = false,
    bool inRange = false,
  }) async {
    checkMap(map);
    return appendColumns(
      [map],
      fromRow: fromRow,
      mapTo: mapTo,
      appendMissing: appendMissing,
      inRange: inRange,
    );
  }

  /// Appends columns with values from [maps].
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [maps] - list of maps containing values to insert (not null nor empty)
  ///
  /// [fromRow] - optional (defaults to 1), row index for the first inserted value,
  /// rows start at index 1
  ///
  /// [mapTo] - optional (defaults to 1), index of a column to which
  /// keys of the [maps] will be mapped to,
  /// columns start at index 1 (column A)
  ///
  /// [appendMissing] - optional (defaults to `false`), whether keys of [maps]
  /// (with its related values) that are not present in a [mapTo] column
  /// should be added
  ///
  /// [inRange] - optional (defaults to false), whether [maps] values should be
  /// appended to last column in range (respects [fromRow]) or last column in table
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> appendColumns(
    List<Map<String, dynamic>> maps, {
    int fromRow = 1,
    int mapTo = 1,
    bool appendMissing = false,
    bool inRange = false,
  }) async {
    check('fromRow', fromRow);
    check('mapTo', mapTo);
    checkMap(maps);
    final columns = await _values.allColumns();
    final column =
        inRange ? inRangeIndex(columns, fromRow) + 1 : columns.length + 1;
    if (column < 2) {
      if (appendMissing && mapTo == 1) {
        return _insertColumns(
          2,
          maps,
          <String>[],
          fromRow,
          mapTo,
          appendMissing,
          false,
        );
      }
      return false;
    }
    checkMapTo(column, mapTo);
    if (mapTo > column) return false;
    final keys = columns[mapTo - 1];
    return _insertColumns(
      column,
      maps,
      keys,
      fromRow,
      mapTo,
      appendMissing,
      false,
    );
  }

  Future<bool> _insertRows(
    int row,
    List<Map<String, dynamic>> maps,
    List<String> allKeys,
    int fromColumn,
    int mapTo,
    bool appendMissing,
    bool overwrite,
  ) async {
    final keys = extractSublist(allKeys, from: fromColumn - 1);
    final newKeys = <String>{};
    final rows = <List>[];
    var newKeysInsertion = Future.value(true);
    if (appendMissing) {
      for (var map in maps) {
        newKeys.addAll(map.keys);
      }
      newKeys.removeAll(allKeys);
      if (newKeys.isNotEmpty) {
        newKeysInsertion = _values.insertRow(
          mapTo,
          newKeys.toList(),
          fromColumn: fromColumn + keys.length,
        );
      }
    }
    if (overwrite) {
      for (var map in maps) {
        final row = [];
        rows.add(row);
        for (var column in keys) {
          row.add(map[column] ?? '');
        }
        for (var column in newKeys) {
          row.add(map[column] ?? '');
        }
      }
    } else {
      for (var map in maps) {
        final row = [];
        rows.add(row);
        for (var column in keys) {
          row.add(map[column]);
        }
        for (var column in newKeys) {
          row.add(map[column]);
        }
      }
    }
    await newKeysInsertion;
    return _values.insertRows(
      row,
      rows,
      fromColumn: fromColumn,
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
    Map<String, dynamic> map, {
    int fromColumn = 1,
    int mapTo = 1,
    bool appendMissing = false,
    bool overwrite = false,
  }) async {
    checkMap(map);
    return insertRows(
      row,
      [map],
      fromColumn: fromColumn,
      mapTo: mapTo,
      appendMissing: appendMissing,
      overwrite: overwrite,
    );
  }

  /// Updates rows with values from [maps].
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [maps] - list of maps containing values to insert (not null nor empty)
  ///
  /// [row] - row index to insert values of [maps] to,
  /// rows start at index 1
  ///
  /// [fromColumn] - optional (defaults to 1), column index for the first inserted
  /// value,
  /// columns start at index 1 (column A)
  ///
  /// [mapTo] - optional (defaults to 1), index of a row to which
  /// keys of the [maps] will be mapped to,
  /// rows start at index 1
  ///
  /// [appendMissing] - optional (defaults to `false`), whether keys of [maps]
  /// (with its related values) that are not present in a [mapTo] row
  /// should be added
  ///
  /// [overwrite] - optional (defaults to `false`), whether clear cells of
  /// [row] if map of [maps] does not contain value for them
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insertRows(
    int row,
    List<Map<String, dynamic>> maps, {
    int fromColumn = 1,
    int mapTo = 1,
    bool appendMissing = false,
    bool overwrite = false,
  }) async {
    check('row', row);
    checkMapTo(row, mapTo);
    checkMap(maps);
    final keys = await _values.row(mapTo);
    return _insertRows(
      row,
      maps,
      keys,
      fromColumn,
      mapTo,
      appendMissing,
      overwrite,
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
  /// [eager] - optional (defaults to `true`), whether to add [key] if absent
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insertRowByKey(
    dynamic key,
    Map<String, dynamic> map, {
    int fromColumn = 2,
    dynamic mapTo,
    bool appendMissing = false,
    bool overwrite = false,
    bool eager = true,
  }) async {
    final rKey = parseKey(key);
    final mKey = parseMapToKey(mapTo);
    check('fromColumn', fromColumn);
    checkMapTo(rKey, mKey);
    checkMap(map);
    final rows = await _values.allRows();
    final mapToIndex = isNullOrEmpty(mKey) ? 0 : whereFirst(rows, mKey);
    if (mapToIndex < 0) return false;
    var rowIndex = whereFirst(rows, rKey);
    if (rowIndex < 0) {
      if (!eager || rows.isEmpty || rows.first.isEmpty) return false;
      rowIndex = rows.length;
      await _values._ws._update(
        values: [rKey],
        range: await _values._ws._columnRange(1, rowIndex + 1, 1),
        majorDimension: DIMEN_COLUMNS,
      );
    } else {
      checkMapTo(rowIndex + 1, mapToIndex + 1);
    }
    final keys = get(rows, at: mapToIndex, or: <String>[]);
    return _insertRows(
      rowIndex + 1,
      [map],
      keys,
      fromColumn,
      mapToIndex + 1,
      appendMissing,
      overwrite,
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
  /// [inRange] - optional (defaults to false), whether [map] values should be
  /// appended to last row in range (respects [fromColumn]) or last row in table
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> appendRow(
    Map<String, dynamic> map, {
    int fromColumn = 1,
    int mapTo = 1,
    bool appendMissing = false,
    bool inRange = false,
  }) async {
    checkMap(map);
    return appendRows(
      [map],
      fromColumn: fromColumn,
      mapTo: mapTo,
      appendMissing: appendMissing,
      inRange: inRange,
    );
  }

  /// Appends rows with values from [maps].
  ///
  /// Expands current sheet's size if inserting range is out of sheet's bounds.
  ///
  /// [maps] - list of maps containing values to insert (not null nor empty)
  ///
  /// [fromColumn] - optional (defaults to 1), column index for the first inserted
  /// value,
  /// columns start at index 1 (column A)
  ///
  /// [mapTo] - optional (defaults to 1), index of a row to which
  /// keys of the [maps] will be mapped to,
  /// rows start at index 1
  ///
  /// [appendMissing] - optional (defaults to `false`), whether keys of [maps]
  /// (with its related values) that are not present in a [mapTo] row
  /// should be added
  ///
  /// [inRange] - optional (defaults to false), whether [maps] values should be
  /// appended to last row in range (respects [fromColumn]) or last row in table
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> appendRows(
    List<Map<String, dynamic>> maps, {
    int fromColumn = 1,
    int mapTo = 1,
    bool appendMissing = false,
    bool inRange = false,
  }) async {
    check('fromColumn', fromColumn);
    check('mapTo', mapTo);
    checkMap(maps);
    final rows = await _values.allRows();
    final row = inRange ? inRangeIndex(rows, fromColumn) + 1 : rows.length + 1;
    if (row < 2) {
      if (appendMissing && mapTo == 1) {
        return _insertRows(
          2,
          maps,
          <String>[],
          fromColumn,
          mapTo,
          appendMissing,
          false,
        );
      }
      return false;
    }
    checkMapTo(column, mapTo);
    if (mapTo > row) return false;
    final keys = rows[mapTo - 1];
    return _insertRows(
      row,
      maps,
      keys,
      fromColumn,
      mapTo,
      appendMissing,
      false,
    );
  }
}

/// Representation of a [Cell].
class Cell implements Comparable {
  final Worksheet _ws;
  final int row;
  final int column;
  String _label;
  String value;
  final bool _insertable;

  Cell._(
    this._ws,
    this.row,
    this.column,
    this.value, [
    this._insertable = true,
  ]);

  /// Returns position of a cell in A1 notation.
  String get label {
    _label ??= '${getColumnLetter(column)}${row}';
    return _label;
  }

  String get worksheetTitle => _ws._title;

  /// Updates value of a cell.
  ///
  /// Returns Future `true` in case of success
  ///
  /// Throws [GSheetsException].
  Future<bool> post(dynamic value) async {
    final val = parseValue(value);
    if (this.value == val) return false;
    final posted = await _ws._update(
      values: [val ?? ''],
      range: "'$worksheetTitle'!$label",
      majorDimension: DIMEN_COLUMNS,
    );
    if (posted) this.value = val;
    return posted;
  }

  /// Refreshes value of a cell.
  ///
  /// Returns Future `true` if value has been changed.
  ///
  /// Throws [GSheetsException].
  Future<bool> refresh() async {
    final before = value;
    final range = "'$worksheetTitle'!$label:$label";
    value = getOrEmpty(await _ws._get(range, DIMEN_COLUMNS));
    return before != value;
  }

  @override
  String toString() => "'$value' at $label";

  @override
  int compareTo(other) {
    return row + column - other.row - other.column;
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

  List<Cell> _wrapColumn(
    List<String> list,
    int column,
    int fromRow,
  ) {
    var row = fromRow;
    return List.unmodifiable(
      list.map((value) => Cell._(_ws, row++, column, value)),
    );
  }

  List<Cell> _wrapRow(
    List<String> list,
    int row,
    int fromColumn,
  ) {
    var column = fromColumn;
    return List.unmodifiable(
      list.map((value) => Cell._(_ws, row, column++, value)),
    );
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
    final list = await _ws.values.column(
      column,
      fromRow: fromRow,
      length: length,
    );
    return _wrapColumn(list, column, fromRow);
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
    final list = await _ws.values.row(
      row,
      fromColumn: fromColumn,
      length: length,
    );
    return _wrapRow(list, row, fromColumn);
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
    dynamic key, {
    int fromRow = 2,
    int length = -1,
  }) async {
    final cKey = parseKey(key);
    check('fromRow', fromRow);
    final columns = await _ws.values.allColumns();
    final columnIndex = whereFirst(columns, cKey);
    if (columnIndex < 0) return null;
    final list = extractSublist(
      columns[columnIndex],
      from: fromRow - 1,
      length: length,
    );
    return _wrapColumn(list, columnIndex + 1, fromRow);
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
    dynamic key, {
    int fromColumn = 2,
    int length = -1,
  }) async {
    final rKey = parseKey(key);
    check('fromColumn', fromColumn);
    final rows = await _ws.values.allRows();
    final rowIndex = whereFirst(rows, rKey);
    if (rowIndex < 0) return null;
    final list = extractSublist(
      rows[rowIndex],
      from: fromColumn - 1,
      length: length,
    );
    return _wrapRow(list, rowIndex + 1, fromColumn);
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
  /// [inRange] - optional (defaults to false), whether should be fetched last
  /// column in range (respects [fromRow] and [length]) or last column in table
  ///
  /// Returns last column as Future [List] of [Cell].
  /// Returns Future `null` if there are no columns.
  ///
  /// Throws [GSheetsException].
  Future<List<Cell>> lastColumn({
    int fromRow = 1,
    int length = -1,
    bool inRange = false,
  }) async {
    if (inRange) {
      final columns = await _ws.values.allColumns(
        fromRow: fromRow,
        length: length,
      );
      if (columns.isEmpty) return null;
      return _wrapColumn(columns.last, columns.length, fromRow);
    } else {
      check('fromRow', fromRow);
      final columns = await _ws.values.allColumns();
      if (columns.isEmpty) return null;
      final list = extractSublist(
        columns.last,
        from: fromRow - 1,
        length: length,
      );
      return _wrapColumn(list, columns.length, fromRow);
    }
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
  /// [inRange] - optional (defaults to false), whether should be fetched last
  /// row in range (respects [fromColumn] and [length]) or last row in table
  ///
  /// Returns last row as Future [List] of [Cell].
  /// Returns Future `null` if there are no rows.
  ///
  /// Throws [GSheetsException].
  Future<List<Cell>> lastRow({
    int fromColumn = 1,
    int length = -1,
    bool inRange = false,
  }) async {
    if (inRange) {
      final rows = await _ws.values.allRows(
        fromColumn: fromColumn,
        length: length,
      );
      if (rows.isEmpty) return null;
      return _wrapRow(rows.last, rows.length, fromColumn);
    } else {
      check('fromColumn', fromColumn);
      final rows = await _ws.values.allRows();
      if (rows.isEmpty) return null;
      final list = extractSublist(
        rows.last,
        from: fromColumn - 1,
        length: length,
      );
      return _wrapRow(list, rows.length, fromColumn);
    }
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
  /// [count] - optional (defaults to -1), the number of requested columns
  /// if count is `-1`, all columns starting from [fromColumn] will be returned
  ///
  /// Returns all columns as Future [List] of [List].
  ///
  /// Throws [GSheetsException].
  Future<List<List<Cell>>> allColumns({
    int fromColumn = 1,
    int fromRow = 1,
    int length = -1,
    int count = -1,
  }) async {
    final columns = await _ws.values.allColumns(
      fromColumn: fromColumn,
      fromRow: fromRow,
      length: length,
      count: count,
    );
    return List<List<Cell>>.generate(
      columns.length,
      (index) => _wrapColumn(columns[index], fromColumn + index, fromRow),
    );
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
  /// [count] - optional (defaults to -1), the number of requested rows
  /// if count is `-1`, all rows starting from [fromRow] will be returned
  ///
  /// Returns all rows as Future [List] of [List].
  ///
  /// Throws [GSheetsException].
  Future<List<List<Cell>>> allRows({
    int fromRow = 1,
    int fromColumn = 1,
    int length = -1,
    int count = -1,
  }) async {
    final rows = await _ws.values.allRows(
      fromRow: fromRow,
      fromColumn: fromColumn,
      length: length,
      count: count,
    );
    return List<List<Cell>>.generate(
      rows.length,
      (index) => _wrapRow(rows[index], fromRow + index, fromColumn),
    );
  }

  /// Find cells by value.
  ///
  /// These cells cannot be updated by [WorksheetAsCells] `insert` method,
  /// [Cell]'s post method instead.
  ///
  /// [value] - value to look for
  ///
  /// [fromRow] - optional (defaults to 1), index of a first row in which
  /// [value] looked for (rows before [fromRow] will be skipped),
  /// rows start at index 1
  ///
  /// [fromColumn] - optional (defaults to 1), index of a first column in which
  /// [value] looked for (columns before [fromColumn] will be skipped),
  /// columns start at index 1 (column A)
  ///
  /// [length] - optional (defaults to -1), the length of a table area in which
  /// [value] looked for, starting from [fromRow] and [fromColumn]
  /// if length is `-1`, all rows after [fromRow] and all columns after
  /// [fromColumn] will be used
  ///
  /// Returns cells as Future [List].
  ///
  /// Throws [GSheetsException].
  Future<List<Cell>> findByValue(
    dynamic value, {
    int fromRow = 1,
    int fromColumn = 1,
    int length = -1,
  }) async {
    final valueString = parseValue(value);
    final cells = <Cell>[];
    var rows = await _ws.values.allRows(
      fromRow: fromRow,
      fromColumn: fromColumn,
      length: length,
    );
    if (length > 0) rows = rows.take(length).toList();
    var rowNumber = fromRow;
    for (var row in rows) {
      var colNumber = fromColumn;
      for (var val in row) {
        if (val == valueString) {
          cells.add(Cell._(_ws, rowNumber, colNumber, val, false));
        }
        colNumber++;
      }
      rowNumber++;
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
    return Cell._(_ws, row, column, value, false);
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
    @required dynamic rowKey,
    @required dynamic columnKey,
  }) async {
    final rKey = parseKey(rowKey, 'row');
    final cKey = parseKey(columnKey, 'column');
    final rows = await _ws.values.allRows();
    if (rows.isEmpty) return null;
    final columnIndex = rows.first.indexOf(cKey);
    if (columnIndex < 0) return null;
    final rowIndex = whereFirst(rows, rKey);
    if (rowIndex < 0) return null;
    final value = getOrEmpty(rows[rowIndex], columnIndex);
    return Cell._(_ws, rowIndex + 1, columnIndex + 1, value, false);
  }

  /// Updates cells with values of [values].
  ///
  /// [values] - cells with values to insert (not null nor empty)
  ///
  /// Returns Future `true` in case of success.
  ///
  /// Throws [GSheetsException].
  Future<bool> insert(List<Cell> values) async {
    checkValues(values);
    except(
      !values.first._insertable,
      'Cells returned by findByValue, cell or cellByKeys cannot be inserted, '
      'use Cell\'s post method instead.',
    );
    final range =
        "'${values.first.worksheetTitle}'!${values.first.label}:${values.last.label}";
    final dimen =
        values.first.row == values.last.row ? DIMEN_ROWS : DIMEN_COLUMNS;
    return _ws._update(
      values: values.map((cell) => cell.value).toList(),
      range: range,
      majorDimension: dimen,
    );
  }
}

/// Mapper for [Worksheet]'s cells.
class CellsMapper {
  final WorksheetAsCells _cells;

  CellsMapper._(this._cells);

  Map<String, Cell> _wrapRow(
    List<String> keys,
    List<String> values,
    int row,
    int fromColumn,
  ) {
    final map = mapKeysToValues(
      keys,
      values,
      (index, val) => Cell._(_cells._ws, row, fromColumn + index, val ?? ''),
    );
    return Map.unmodifiable(map);
  }

  Map<String, Cell> _wrapColumn(
    List<String> keys,
    List<String> values,
    int column,
    int fromRow,
  ) {
    final map = mapKeysToValues(
      keys,
      values,
      (index, val) => Cell._(_cells._ws, fromRow + index, column, val ?? ''),
    );
    return Map.unmodifiable(map);
  }

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
    int mapTo = 1,
    int fromRow = 1,
    int length = -1,
  }) async {
    check('column', column);
    check('mapTo', mapTo);
    checkMapTo(column, mapTo);
    final columns = await _cells._ws.values.allColumns(
      fromRow: fromRow,
      length: length,
    );
    final keys = get(columns, at: mapTo - 1, or: <String>[]);
    final values = get(columns, at: column - 1, or: <String>[]);
    return _wrapColumn(keys, values, column, fromRow);
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
    check('row', row);
    check('mapTo', mapTo);
    checkMapTo(row, mapTo);
    final rows = await _cells._ws.values.allRows(
      fromColumn: fromColumn,
      length: length,
    );
    final keys = get(rows, at: mapTo - 1, or: <String>[]);
    final values = get(rows, at: row - 1, or: <String>[]);
    return _wrapRow(keys, values, row, fromColumn);
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
    dynamic key, {
    int fromRow = 2,
    dynamic mapTo,
    int length = -1,
  }) async {
    final cKey = parseKey(key);
    final mKey = parseMapToKey(mapTo);
    check('fromRow', fromRow);
    checkMapTo(cKey, mKey);
    final columns = await _cells._ws.values.allColumns();
    final columnIndex = whereFirst(columns, cKey);
    if (columnIndex < 0) return null;
    final mapToIndex = isNullOrEmpty(mKey) ? 0 : whereFirst(columns, mKey);
    if (mapToIndex < 0) return null;
    checkMapTo(columnIndex + 1, mapToIndex + 1);
    final keys = extractSublist(
      columns[mapToIndex],
      from: fromRow - 1,
      length: length,
    );
    final values = extractSublist(
      columns[columnIndex],
      from: fromRow - 1,
      length: length,
    );
    return _wrapColumn(keys, values, columnIndex + 1, fromRow);
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
    dynamic key, {
    int fromColumn = 2,
    dynamic mapTo,
    int length = -1,
  }) async {
    final rKey = parseKey(key);
    final mKey = parseMapToKey(mapTo);
    check('fromColumn', fromColumn);
    checkMapTo(key, mKey);
    final rows = await _cells._ws.values.allRows();
    final rowIndex = whereFirst(rows, rKey);
    if (rowIndex < 0) return null;
    final mapToIndex = isNullOrEmpty(mKey) ? 0 : whereFirst(rows, mKey);
    if (mapToIndex < 0) return null;
    checkMapTo(rowIndex + 1, mapToIndex + 1);
    final keys = extractSublist(
      rows[mapToIndex],
      from: fromColumn - 1,
      length: length,
    );
    final values = extractSublist(
      rows[rowIndex],
      from: fromColumn - 1,
      length: length,
    );
    return _wrapRow(keys, values, rowIndex + 1, fromColumn);
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
  /// [inRange] - optional (defaults to false), whether should be fetched last
  /// column in range (respects [fromRow] and [length]) or last column in table
  ///
  /// Returns column as Future [Map] of [String] to [Cell].
  /// Returns Future `null` if there are less than 2 columns.
  ///
  /// Throws [GSheetsException].
  Future<Map<String, Cell>> lastColumn({
    int fromRow = 1,
    int mapTo = 1,
    int length = -1,
    bool inRange = false,
  }) async {
    check('mapTo', mapTo);
    if (inRange) {
      final columns = await _cells._ws.values.allColumns(
        fromRow: fromRow,
        length: length,
      );
      final column = columns.length;
      if (column < 2) return null;
      checkMapTo(column, mapTo);
      final keys = get(columns, at: mapTo - 1, or: <String>[]);
      final values = columns[column - 1];
      return _wrapColumn(keys, values, column, fromRow);
    } else {
      check('fromRow', fromRow);
      final columns = await _cells._ws.values.allColumns();
      final column = columns.length;
      if (column < 2) return null;
      checkMapTo(column, mapTo);
      final keys = extractSublist(
        get(columns, at: mapTo - 1, or: <String>[]),
        from: fromRow - 1,
        length: length,
      );
      final values = extractSublist(
        columns[column - 1],
        from: fromRow - 1,
        length: length,
      );
      return _wrapColumn(keys, values, column, fromRow);
    }
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
  /// [inRange] - optional (defaults to false), whether should be fetched last
  /// row in range (respects [fromColumn] and [length]) or last row in table
  ///
  /// Returns row as Future [Map] of [String] to [Cell].
  /// Returns Future `null` if there are less than 2 rows.
  ///
  /// Throws [GSheetsException].
  Future<Map<String, Cell>> lastRow({
    int fromColumn = 1,
    int mapTo = 1,
    int length = -1,
    bool inRange = false,
  }) async {
    check('mapTo', mapTo);
    if (inRange) {
      final rows = await _cells._ws.values.allRows(
        fromColumn: fromColumn,
        length: length,
      );
      final row = rows.length;
      if (row < 2) return null;
      checkMapTo(row, mapTo);
      final keys = get(rows, at: mapTo - 1, or: <String>[]);
      final values = rows[row - 1];
      return _wrapRow(keys, values, row, fromColumn);
    } else {
      check('fromColumn', fromColumn);
      final rows = await _cells._ws.values.allRows();
      final row = rows.length;
      if (row < 2) return null;
      checkMapTo(row, mapTo);
      final keys = extractSublist(
        get(rows, at: mapTo - 1, or: <String>[]),
        from: fromColumn - 1,
        length: length,
      );
      final values = extractSublist(
        rows[row - 1],
        from: fromColumn - 1,
        length: length,
      );
      return _wrapRow(keys, values, row, fromColumn);
    }
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
