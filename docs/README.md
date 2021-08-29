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
| Loading via sheet position or name         | SheetIndex  or SheetName  | SheetIndex = 0      | Setting both will throw `ArgumentException`
| Skip top rows                     | HeaderRowIndex            | 0         |
| Remove empty rows                 | RemoveEmptyRows           | true      |
| Allow duplicate columns          | AllowDuplicateColumns    | true      | Duplicated column will be loaded as `<column name>_<guid>`
| Limit max columns to load         | MaxColumns                | null      |


# Usage

```csharp
using Ninjanaut.IO;

var datatable = ExcelReader.ToDataTable(@"C:\FooExcel.xlsx");
```

or with options argument (the default settings)

```csharp
using Ninjanaut.IO;

var datatable = ExcelReader.ToDataTable(@"C:\FooExcel.xlsx", new() {
                    Format = ExcelReaderFormat.Xlsx,
                    SheetIndex = 0,
                    SheetName = null,
                    HeaderRowIndex = 0,
                    RemoveEmptyRows = true,
                    AllowDuplicateColumns = true,
                    MaxColumns = null
                });
```

# Notes

DataTable object is suitable for this purpose, because you can easily view the read data directly in Visual Studio for debug purposes, create a collection of entities from it or pass datatable as parameter directly into the SQL server stored procedure.

# Contribution

If you would like to contribute to the project, please send a pull request to the dev branch.
