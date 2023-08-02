# excel-table-export-tool
Export excel data sheets to lua, json, xml, sqlite, mongo

# Notes
- Support xlsx, xls format files.
- Supports file export to lua, json, xml, sqlite, mongo.
- Supported data types include: int, float, bool, json, string, python dict.
- Rows that are hidden in excel are not exported
- Filter blank lines in excel
- Lines that start with @ignore are ignored

# Usage
Columns 1 through 3 are defined as the header
The first column defines the data type (int, float, string, json, python dict)
The second column defines the field name
The third column describes the field
Excel worksheet name is the exported file name, or mongo, sqlite table name

example : 
The first three rows are fixed as the headers.
Starting with row 4 is the data row.
Line 3 of the data row is empty and will be skipped.
Line 4 of the data line begins with @ignore and is ignored.

|     int     |    float    |   string    |    json     |    dict     |
|:-----------:|:-----------:|:-----------:|:-----------:|:-----------:|
| field-name1 | field-name2 | field-name3 | field-name4 | field-name5 |
| Description | Description | Description | Description | Description |
|      1      |     1.2     |     abcd    |   [1,2,3]   |  {'a' : 1}  |
|      1      |     1.2     |     abcd    |{"a" : "abc"}|   (1,2,3)   |
|is empty example    |             |             |             |             |
|   @ignore   |     1.2     |     abcd    |   [1,2,3]   |   (1,2,3)   |

Type fields can be decorated with 'key', 'default', 'ignore'.
@key Indicates that this field is used as a key.
@default Indicates that the default value is given by type when the table is empty.
@ignore Indicates that this column of data is ignored.

|   int@key   |float@default|string@ignore|    json     |    dict     |
|:-----------:|:-----------:|:-----------:|:-----------:|:-----------:|
| field-name1 | field-name2 | field-name3 | field-name4 | field-name5 |
| Description | Description | Description | Description | Description |
|      1      |     1.2     |     abcd    |   [1,2,3]   |  {'a' : 1}  |
|      1      |     1.2     |     abcd    |{"a" : "abc"}|   (1,2,3)   |
|is empty example     |             |             |             |             |
|   @ignore   |     1.2     |     abcd    |   [1,2,3]   |   (1,2,3)   |

# Script parameter description
- m/mode Working mode(lua, json, xml, sqlite, mongo)
- o/output Output information, File export path or database information, example : ./test or test.db
- f/file Excel file
- n/names Excel worksheet, Press ',' to split, example : Sheet1,Sheet2

# Example
```
# Export the data to 'test' in the mongo database and export each excel row as a document
python excelop.py -m mongo -o 192.168.3.147:27017@test:1 -f test.xlsx -n Sheet1

# Export the data to "test" in the mongo database and export the entire excel row as a single document
python excelop.py -m mongo -o 192.168.3.147:27017@test -f test.xlsx -n Sheet1

# Export excel data to sqlite's test.db and export each excel row as a row of sql data
python excelop.py -m sqlite -o test.db@1 -f test.xlsx -n Sheet1

# Export excel data to sqlite's test.db, and the entire excel file is exported as a single line of sql data
python excelop.py -m sqlite -o test.db -f test.xlsx -n Sheet1

# Export excel data to 'Sheet1.lua' in the 'test' folder
python excelop.py -m lua -o test -f test.xlsx -n Sheet1

# Export excel data to 'Sheet1.json' in the 'test' folder
python excelop.py -m json -o test -f test.xlsx -n Sheet1

# Export excel data to 'Sheet1.xml' in the 'test' folder
python excelop.py -m xml -o test -f test.xlsx -n Sheet1

```



















