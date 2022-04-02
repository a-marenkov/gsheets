## 0.4.2

- Fix `SpreadsheetData` parsing

## 0.4.1

- Lower some dependencies bounds
- Fix flutter AOT build

## 0.4.0

- Bump googleapis version to ^8.0.0
- Add `SpreadsheetData` to `Spreadsheet` with `SpreadsheetProperties`, `NamedRanges`, `DeveloperMetadata`, `DataSource` and `DataSourceRefreshSchedule`

## 0.3.2

- Bump googleapis version to 4.0.0
- Added A1Ref utility class
- Added method `export` to `Spreadsheet` (with export formats xlsx, csv and pdf)

## 0.3.1

- Documentation fixes.
- Bump googleapis version to 3.0.0

## 0.3.0

- **Breaking change** Migration to nullsafety progress
- **Breaking change** - ValueRenderOption use snake case
- Added `fill` parameter to `allRows` and `allColumns` to return lists with even lengths (if the value for the cell is
  absent, empty string is appended)

## 0.2.7

- Added constructor with custom client for `GSheets`
- Added `close` method to `GSheets`

## 0.2.6

- Updated dependencies

## 0.2.5

- Added `addFromSpreadsheet` method to `Spreadsheet` that copies worksheet from another spreadsheet;
- Added `copyTo` method to `Worksheet` that copies it to another spreadsheet;
- Refactored `refresh` method of `Spreadsheet` - now it updates current list of worksheets instead of supplying new one;
- Fixed inserting and reading all rows/columns if worksheet's number of rows/columns has been changed by some other
  source;
- Added some documentation to Permission and Spreadsheet.

## 0.2.4+1

- Fixed fetching spreadsheets that contain non-grid worksheets

## 0.2.4

- Added `createSpreadsheet` method to `GSheets` that allows to create new spreadsheet;
- Added `batchUpdate` method to `Spreadsheet` that applies one or more updates to the spreadsheet;
- Exposed `batchUpdate` method of `Gsheets`.

## 0.2.3

- Added encoding to ranges - fixes `Unexpected character exception` if sheet's name contains character `/` or other
  characters that has to be encoded.

## 0.2.1+1

- Readme update - added link to the medium article with tutorial.

## 0.2.1

- **Non-Breaking major change**: keys (for methods that use them) and values (for methods that update values) are
  made `dynamic`;
- added `ValueRenderOption` and `ValueInputOption` to `spreadsheet` method of `GSheets`;
- added methods to `WorksheetAsValues` that insert multiple rows/columns (`insertRows`, `insertColumns`,`appendRows`
  ,`appendColumns`);
- added methods to `ValuesMapper` that insert multiple rows/columns (`insertRows`, `insertColumns`,`appendRows`
  ,`appendColumns`);
- added `eager` to the methods that insert values by keys.

## 0.2.0+1

- Readme update - added link to the article with credentials instructions.

## 0.2.0

- Increased performance of methods that include mapping, use keys, fetch or append last row/column;
- Decreased Google API calls for methods that include mapping, use keys, fetch or append last row/column;
- Added `allColumns` and `allRows` methods to `ValuesMapper`;
- Added `count` parameter to `allRows` and `allColumns` methods of `WorksheetAsCells`;
- Added `inRange` parameter to methods that fetch or append last row/column;
- **Breaking change**: parameter `user` of `Spreadsheet`s method `share` was made positional;
- **Breaking change**: `Cell`'s fields `rowIndex` and `columnIndex` were renamed to `row` and `column`;
- Minor fixes.

## 0.1.1+1

- Refactoring according to dart analysis.

## 0.1.1

- Lowered version of meta package to make gsheets compatible with flutter stable branch.

## 0.1.0

- Initial release;
- Minor fixes.

## 0.1.0-rc.1

- Added `lastRow` and `lastColumn` methods to `CellsMapper`;
- Refined documentation;
- Minor fixes.

## 0.0.1-dev.5+1

- Documentation fixes.

## 0.0.1-dev.5

- Refactoring with breaking changes - `lenght` parameter of some `Worksheet` methods was renamed to `count` as more
  appropriate;
- added method `add` to `Worksheet` that adds new columns and rows;
- Bug fixes.

## 0.0.1-dev.4

- Refactoring with breaking changes - simplified GSheets initialization;
- Increased performance;
- Minor fixes.

## 0.0.1-dev.3

- Documentation fixes.

## 0.0.1-dev.2

- Refactoring with breaking changes. Some methods and its parameters were renamed to increase readability. Some
  additional classes were exposed.

## 0.0.1-dev.1

- Initial version.
