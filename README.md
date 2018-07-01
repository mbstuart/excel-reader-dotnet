# excel-reader-dotnet

A C# library for extracting data from xlsx files 

This uses NPOI in order to read the XLSX file.

### Set up
Dependencies (NPOI) must be downloaded via NuGet

### Project structure
1. ExcelUploadServices contains the implementation code for the excel reader, including some internal interfaces 
2. ExcelUploadServices.Interfaces contains publicly exposed interfaces, in addition to the models being used

### Input excel
Excels must have:
1. One piece of tabular data per sheet
2. Top left cell of the table should be a named cell with the following format: "TableId_<TableIdName>"
3. Top row is assumed to be a columns row
3. The data reader terminates once the leftmost column is left blank

### Output 
Returns the following data:
1. A list of columns, with the name provided being used as the column name
2. A list of rows (with a row being represented as a List of Objects with the index of the object corresponding to the index within the list of columns, so the 4th item in row x is referenced by the 4th column)


