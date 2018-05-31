# Json to Excel

To whom it may concern, but remember HKNOIT

## Getting Started

Easy convert resultset to json and excel

### Library

* [Unirest](http://unirest.io/) - http request library
* [Apache POI](https://poi.apache.org/) - Java Excel API


## Sample Code

Resultset to Json String

```
String json = ResultSetToJson.ResultSetToJsonString(resultSet);
```

Json Object to Excel

```
HttpResponse<JsonNode> getJsonNode = Unirest.get("https://yourapi.com/getsomething").asJson();

JsonObjectToExcel jsonObjectToExcel = new JsonObjectToExcel();
HSSFWorkbook hssfWorkbook = jsonObjectToExcel.dump(getJsonNode);
FileOutputStream fileOutputStream = new FileOutputStream(outputFile);
hssfWorkbook.write(fileOutputStream);
hssfWorkbook.close();
```

Resultset to Excel

```
ResultSetToExcel resultSetToExcel = new ResultSetToExcel();
LinkedHashMap<String, String> linkedHashMap = new LinkedHashMap<String, String>();

/* custom number format */
linkedHashMap.put("float.*","0.00");
linkedHashMap.put("decimal.*","0.00");

resultSetToExcel.setMapSqlTypeExcelFormat(linkedHashMap);

HSSFWorkbook workbook = resultSetToExcel.dump(resultset);

FileOutputStream fileOutputStream = new FileOutputStream(outputFile);
workbook.write(fileOutputStream);
workbook.close();
```