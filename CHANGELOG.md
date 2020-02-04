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

- Refactoring with breaking changes - `lenght` parameter of some `Worksheet` methods was renamed to `count` as more appropriate;
- added method `add` to `Worksheet` that adds new columns and rows;
- Bug fixes.

## 0.0.1-dev.4

- Refactoring with breaking changes - simplified GSheets initialization;
- Increased performance;
- Minor fixes.

## 0.0.1-dev.3

- Documentation fixes.

## 0.0.1-dev.2

- Refactoring with breaking changes. Some methods and its parameters were renamed 
to increase readability. Some additional classes were exposed.

## 0.0.1-dev.1

- Initial version.
