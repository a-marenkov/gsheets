# Gsheets with Dart :page_facing_up:

A library for working with Google Sheets API v4.

Manage your spreadsheets with gsheets in Dart.

## Usage :wrench:

Basic usage see [example][example]. For more advanced examples, check out following article [Dart: Working With Google Sheets][tutorial].

If you don't know where to find the credentials, i recommend you to read following article [How To Get Credentials for Google Sheets][credentials].

1. After you setup the credentials for your API, you have to download a json from the google console that is like the following:

```json
{
  "type": "service_account",
  "project_id": "",
  "private_key_id": "",
  "private_key": "",
  "client_email": "",
  "client_id": "",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": ""
}
```

You have to insert that json content in the **\_credentials** variable,so it has to be like this:

```
const _credentials = r'''
{
  "type": "service_account",
  "project_id": "",
  "private_key_id": "",
  "private_key": "",
  "client_email": "",
  "client_id": "",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": ""
}
''';
```

2. Add the spreadsheet id, you can find it in the url of your google sheets document:

```
https://docs.google.com/spreadsheets/d/spreadsheetId/edit#gid=1590950017
```

so you have to paste the **spreadsheetId** in the **\_spreadsheetId** variable:

```dart
const _spreadsheetId = 'spreadsheetId';
```

3. then you can use the library, here we have some examples, you can check more commands in the [example] folder

```dart

void main() async {
  // init GSheets
  final gsheets = GSheets(_credentials);
  // fetch spreadsheet by its id
  final ss = await gsheets.spreadsheet(_spreadsheetId);
  // get worksheet by its title
  var sheet =  ss.worksheetByTitle('example');
  // create worksheet if it does not exist yet
  sheet ??= await ss.addWorksheet('example');

  // update cell at 'B2' by inserting string 'new'
  await sheet.values.insertValue('new', column: 2, row: 2);
  // prints 'new'
  print(await sheet.values.value(column: 2, row: 2));
  // get cell at 'B2' as Cell object
  final cell = await sheet.cells.cell(column: 2, row: 2);
  // prints 'new'
  print(cell.value);
  // update cell at 'B2' by inserting 'new2'
  await cell.post('new2');
  // prints 'new2'
  print(cell.value);
  // also prints 'new2'
  print(await sheet.values.value(column: 2, row: 2));
 // append row in the followind columns with his values
   final secondRow = {
    'index': '5',
    'letter': 'f',
    'number': '6',
    'label': 'f6',
  };
  await sheet.values.map.appendRow(secondRow);
  // prints {index: 5, letter: f, number: 6, label: f6}
  print(await sheet.values.map.lastRow());
}
```

4. Of course, you can use with flutter, you can check the [example_gsheets] folder, basically you can read data and also write in the sheet, here the screenshots.
<p float="left">
<img src="https://github.com/WilliBobadilla/gsheets/blob/example_flutter/media/screenshoots/1.jpeg"  width="25%" height="35%" />
<img src="https://github.com/WilliBobadilla/gsheets/blob/example_flutter/media/screenshoots/2.jpeg"  width="25%" height="35%" />
<img src="https://github.com/WilliBobadilla/gsheets/blob/example_flutter/media/screenshoots/3.jpeg"  width="25%" height="35%" />
<img src="https://github.com/WilliBobadilla/gsheets/blob/example_flutter/media/screenshoots/4.jpeg"  width="25%" height="35%" />
</p>

## Upcoming :rocket:

**0.3.\_** - will be added class Table that holds mutable rows and columns and has a sync method.

## Features and bugs :beetle:

Please file feature requests and bugs at the [issue tracker][tracker].

[tracker]: https://github.com/a-marenkov/gsheets/issues
[example]: https://github.com/a-marenkov/gsheets/tree/master/example
[credentials]: https://medium.com/@a.marenkov/how-to-get-credentials-for-google-sheets-456b7e88c430
[tutorial]: https://medium.com/@a.marenkov/dart-working-with-google-sheets-793ed322daa0
[example_gsheets]: https://github.com/WilliBobadilla/gsheets/tree/example_flutter/example_gsheets
