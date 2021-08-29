Welcome to the ExcelReader project website! ExcelReader is lightweight C# library to ease 
loading data from Excel file into DataTable object, based on NPOI library.

# Installation

from nuget package manager console
```powershell
PM> Install-Package Ninjanaut.ExcelReader
```
from command line
```cmd
> dotnet add package Ninjanaut.ExcelReader
```

| Version | Targets |
|- |- |
| 1.x | .NET 5 |

# Features

|   | Options   | Defaults  | Notes |
| -                                 | -                         | -         | - |
| Supports xsl, xlsx and xlsm formats   | Format                    | xlsx      |
| Loading via sheet position or name         | SheetIndex  or SheetName  | SheetIndex = 0      |
| Skip top rows                     | HeaderRowIndex            | 0         |
| Remove empty rows                 | RemoveEmptyRows           | true      |
| Allow duplicate columns          | AllowDuplicateColumns    | true      | Duplicated column will be loaded as `<column name>_<guid>`
| Limit max columns to load         | MaxColumns                | null      |


# Usage

```csharp
using Ninjanaut.IO;

// From file path
var path = @"C:\FooExcel.xlsx";
var datatable = ExcelReader.ToDataTable(path);

// Or from bytes
var path = @"C:\FooExcel.xlsx";
var bytes = File.ReadAllBytes(path);
var datatable = ExcelReader.ToDataTable(bytes);
```

you can also use options argument

```csharp
using Ninjanaut.IO;

var path = @"C:\FooExcel.xlsx";
var options = new ExcelReaderOptions 
{ 
    // Default settings:
    Format = ExcelReaderFormat.Xlsx,
    SheetIndex = 0,
    SheetName = null,
    HeaderRowIndex = 0,
    RemoveEmptyRows = true,
    AllowDuplicateColumns = true,
    MaxColumns = null
});

var datatable = ExcelReader.ToDataTable(path, options);

// The options can be defined within the method.
var datatable = ExcelReader.ToDataTable(path, new() { SheetName = "My Sheet" });
```

# Notes

DataTable object is suitable for this purpose, because you can easily view the read data directly in Visual Studio for debug purposes, create a collection of entities from it or pass datatable as parameter directly into the SQL server stored procedure.

# Contribution

If you would like to contribute to the project, please send a pull request to the dev branch.
