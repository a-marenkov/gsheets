//import 'package:example_gsheets/dialog_custom.dart';
import 'package:flutter/gestures.dart';
import 'package:flutter/material.dart';
import 'package:example_gsheets/models/product_model.dart';
import 'package:gsheets/gsheets.dart';
import 'package:example_gsheets/constants.dart';

const credentials = r'''
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
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/gssheets%40parmalatpilarapi.iam.gserviceaccount.com"
}
''';

const spreadsheetId = '1ZF6pmoreletterandnumbersmjqkb0IOw';
void main() {
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'Flutter Demo',
      theme: ThemeData(
        primarySwatch: Colors.blue,
        visualDensity: VisualDensity.adaptivePlatformDensity,
      ),
      home: MyHomePage(title: 'Flutter Gsheets Demo Page'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  MyHomePage({Key key, this.title}) : super(key: key);

  final String title;

  @override
  _MyHomePageState createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  TextEditingController textSearchController = TextEditingController();
  TextEditingController inputNameController = TextEditingController();
  TextEditingController inputQtyController = TextEditingController();
  TextEditingController inputCapacityController = TextEditingController();
  List<ProductModel> dataFromSheet = List<ProductModel>();
  List<ProductModel> filteredList = List<ProductModel>();
  final gsheets = GSheets(credentials);

  // fetch spreadsheet by its id
  var ss;
  var sheet;
  bool filtered = false;
  bool searching = false;
  bool loading = false;
  _MyHomePageState() {
    onInit();
  }
  void onInit() async {
    print("ejecutando");
    // setState(() {
    loading = true;
    //});

    spreadSheetsReading();
    Map data = await spreadSheetsReading();
    print(data);
    dataFromSheet.clear();
    setState(() {
      for (int i = 0; i < data['name'].length; i++) {
        dataFromSheet.add(ProductModel(
            name: data['name'][i],
            size: int.parse(data['capacity'][i]),
            amount: int.parse(data['quantity'][i]),
            category: ' '));
      }
      loading = false;
    });
  }

  void deleteFilter() {
    setState(() {
      filtered = false;
      textSearchController.clear();
    });
  }

  void showAnimation(bool activated) {
    setState(() {
      searching = activated;
    });
  }

  void onItemChanged(String query) {
    searchResults(query);
    setState(() {
      if (!query.isNotEmpty) {
        filtered = false;
      }
    });
  }

  void searchResults(String query) {
    filtered = true;
    filteredList.clear();
    print("datos" + filtered.toString());
    setState(() {
      dataFromSheet.forEach((element) {
        var lowerNameProduct = element.name.toLowerCase();
        if (lowerNameProduct.contains(query)) {
          filteredList.add(element);
        }
        print(element.name.toString().toLowerCase());
      });
    });

    print("--------------------filtrado--------------");
    print(filteredList.toString());
  }

  Future<Map> spreadSheetsReading() async {
    ss = await gsheets.spreadsheet(spreadsheetId);

    // get worksheet by its title
    sheet = ss.worksheetByTitle('example');
    // create worksheet if it does not exist yet
    sheet ??= await ss.addWorksheet('example');

    var name = await sheet.values.columnByKey("name");
    var quantity = await sheet.values.columnByKey("quantity");
    var capacity = await sheet.values.columnByKey("capacity");

    return {'name': name, 'quantity': quantity, 'capacity': capacity};
  }

  void insertData() async {
    print("insertando");
    print(inputNameController.text);
    print(inputQtyController.text);
    print(inputCapacityController.text);
    final row = {
      'name': inputNameController.text,
      'quantity': inputQtyController.text,
      'capacity': inputCapacityController.text,
    };
    await sheet.values.map.appendRow(row);
    onInit(); //to update the list
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        actions: <Widget>[
          filtered
              ? IconButton(
                  icon: Icon(Icons.delete),
                  tooltip: 'Filtered',
                  onPressed: () {
                    deleteFilter();
                  },
                )
              : Container(),
          searching
              ? IconButton(
                  icon: Icon(Icons.clear),
                  tooltip: 'Clear',
                  onPressed: () {
                    showAnimation(false);
                    // showSearch(context: context, delegate: CustomSearchDelegate());
                  },
                )
              : IconButton(
                  icon: Icon(Icons.search),
                  tooltip: 'Search',
                  onPressed: () {
                    showAnimation(true);
                    // showSearch(context: context, delegate: CustomSearchDelegate());
                  },
                )
        ],
        title: Text('Example Gsheets'),
        centerTitle: true,
      ),
      body: body(),
      floatingActionButton: FloatingActionButton(
          child: Icon(Icons.add),
          onPressed: () {
            //spreadSheetsReading();
            showDialog(
                context: context,
                builder: (BuildContext context) {
                  return customDialogBox();
                });
          }),
    );
  }

  Widget body() {
    return Stack(children: [
      SingleChildScrollView(
          dragStartBehavior: DragStartBehavior.down,
          child: Column(children: [
            AnimatedContainer(
                duration: Duration(milliseconds: 800),
                curve: Curves.fastOutSlowIn,
                height: searching ? 60 : 0,
                child: Visibility(
                    visible: searching,
                    child: Padding(
                      padding: const EdgeInsets.all(12.0),
                      child: TextField(
                        controller: textSearchController,
                        decoration: InputDecoration(
                          hintText: 'Search...',
                        ),
                        onChanged: (String value) {
                          onItemChanged(value);
                        },
                      ),
                    ))),
            Visibility(
                visible: loading,
                child: CircularProgressIndicator(
                  strokeWidth: 10,
                  backgroundColor: Colors.cyanAccent,
                  valueColor: new AlwaysStoppedAnimation<Color>(Colors.red),
                  // value: _progress,
                )),
            ListView.builder(
                physics: NeverScrollableScrollPhysics(),
                scrollDirection: Axis.vertical,
                shrinkWrap: true,
                itemCount:
                    !filtered ? dataFromSheet.length : filteredList.length,
                itemBuilder: (BuildContext context, int index) {
                  return personalizedCard(index);
                }),
          ]))
    ]);
  }

  Widget personalizedCard(int index) {
    return Container(
        child: Card(
            elevation: 2,
            child: ListTile(
              title: !filtered
                  ? Column(
                      children: [
                        Row(
                            mainAxisAlignment: MainAxisAlignment.spaceAround,
                            children: [
                              Text(dataFromSheet[index].name),
                              Text(
                                "Capacity ml: " +
                                    dataFromSheet[index].size.toString(),
                              )
                            ]),
                        Row(
                            mainAxisAlignment: MainAxisAlignment.end,
                            children: [
                              Text("Quantity:  " +
                                  dataFromSheet[index].amount.toString())
                            ])
                      ],
                    )
                  : Column(
                      children: [
                        Row(
                            mainAxisAlignment: MainAxisAlignment.spaceAround,
                            children: [
                              Text(filteredList[index].name),
                              Text("Capacity ml: " +
                                  filteredList[index].size.toString())
                            ]),
                        Row(
                            mainAxisAlignment: MainAxisAlignment.end,
                            children: [
                              Text("Quantity:  " +
                                  filteredList[index].amount.toString())
                            ])
                      ],
                    ),
              leading: Icon(Icons.local_drink),
            )));
  }

  Widget customDialogBox() {
    return Dialog(
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(Constants.padding),
      ),
      elevation: 0,
      backgroundColor: Colors.transparent,
      child: Stack(
        children: <Widget>[
          Container(
            padding: EdgeInsets.only(
                left: Constants.padding,
                top: Constants.avatarRadius + Constants.padding,
                right: Constants.padding,
                bottom: Constants.padding),
            margin: EdgeInsets.only(top: Constants.avatarRadius),
            decoration: BoxDecoration(
                shape: BoxShape.rectangle,
                color: Colors.white,
                borderRadius: BorderRadius.circular(Constants.padding),
                boxShadow: [
                  BoxShadow(
                      color: Colors.black,
                      offset: Offset(0, 10),
                      blurRadius: 10),
                ]),
            child: Column(
              mainAxisSize: MainAxisSize.min,
              children: <Widget>[
                Text(
                  "Insert data",
                  style: TextStyle(fontSize: 22, fontWeight: FontWeight.w600),
                ),
                SizedBox(
                  height: 15,
                ),
                Column(children: [
                  ListTile(
                      leading: Text("Name:"),
                      title: TextField(
                        controller: inputNameController,
                        decoration: InputDecoration(
                          hintText: '..',
                        ),
                        onChanged: (String value) {
                          // onItemChanged(value);
                        },
                      )),
                  ListTile(
                      leading: Text("Quantity:"),
                      title: TextField(
                        controller: inputQtyController,
                        keyboardType: TextInputType.number,
                        decoration: InputDecoration(
                          hintText: '..',
                        ),
                        onChanged: (String value) {
                          //onItemChanged(value);
                        },
                      )),
                  ListTile(
                      leading: Text("Capacity:"),
                      title: TextField(
                        controller: inputCapacityController,
                        keyboardType: TextInputType.number,
                        decoration: InputDecoration(
                          hintText: 'ml',
                        ),
                        onChanged: (String value) {
                          //onItemChanged(value);
                        },
                      )),
                ]),
                SizedBox(
                  height: 22,
                ),
                Row(
                    mainAxisAlignment: MainAxisAlignment.spaceAround,
                    children: [
                      FlatButton(
                          onPressed: () {
                            Navigator.of(context).pop();
                          },
                          child: Text(
                            "Cancel",
                            style: TextStyle(fontSize: 18),
                          )),
                      FlatButton(
                          onPressed: () {
                            insertData();
                            Navigator.of(context).pop();
                          },
                          child: Text(
                            "Acept",
                            style: TextStyle(fontSize: 18),
                          ))
                    ]),
              ],
            ),
          ),
          Positioned(
            left: Constants.padding,
            right: Constants.padding,
            child: CircleAvatar(
              backgroundColor: Colors.transparent,
              radius: Constants.avatarRadius,
              /*  child: ClipRRect(
                  borderRadius:
                      BorderRadius.all(Radius.circular(Constants.avatarRadius)),
                  child: Image.asset("assets/model.jpeg")),*/
            ),
          ),
        ],
      ),
    );
  }
}
