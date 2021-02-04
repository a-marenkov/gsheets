class ProductModel {
  String name;
  int size;
  int amount;
  String category;

  ProductModel({this.name, this.size, this.amount, this.category});
  @override
  String toString() {
    return "{" +
        this.name +
        " " +
        this.size.toString() +
        " " +
        this.amount.toString() +
        " " +
        this.category +
        " " +
        "}";
  }
  //we have to add toJson()
}
