using ExcelReaderTests.Utilities;
using Ninjanaut.IO;
using System;
using System.Data;
using System.Globalization;
using Xunit;

namespace Ninjanaut.ExcelReaderTests
{
    public class ExcelReaderTests
    {
        [Fact]
        public void Load_excel_and_remove_empty_header_rows()
        {
            // Act
            var datatable = ExcelReader.ToDataTable(@"TestData\EmptyRowsFromTop.xlsx", new() { HeaderRowIndex = 2 });

            // Assert
            var dt = new DataTable();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("C")
            });

            dt.AddRow(new object[] { "1", "2", "3"});
            dt.AddRow(new object[] { "1", "2", "3"});
            dt.AddRow(new object[] { "1", "2", "3"});

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.DataTables(datatable, dt);
        }

        [Fact]
        public void Load_excel_and_remove_empty_rows()
        {
            // Act
            var datatable = ExcelReader.ToDataTable(@"TestData\EmptyRows.xlsx", new() { RemoveEmptyRows = true });

            // Assert
            var dt = new DataTable();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("C")
            });

            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.DataTables(datatable, dt);
        }

        [Fact]
        public void Load_excel_with_empty_rows()
        {
            // Act
            // Each empty row contains formatting so that excel also knows that it is the row to include.
            var datatable = ExcelReader.ToDataTable(@"TestData\EmptyRows.xlsx", new() { RemoveEmptyRows = false });

            // Assert
            var dt = new DataTable();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("C")
            });

            dt.AddRow(new object[] { "", "", "" });
            dt.AddRow(new object[] { "", "", "" });
            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "", "", "" });
            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "", "", "" });
            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "", "", "" });
            dt.AddRow(new object[] { "", "", "" });

            Assert.NotNull(datatable);
            Assert.Equal(9, datatable.Rows.Count);
            DataTableAssert.DataTables(datatable, dt);
        }

        [Fact]
        public void Load_excel_with_max_columns_option()
        {
            // Act
            var datatable = ExcelReader.ToDataTable(@"TestData\AdditionalColumns.xlsx", new() { MaxColumns = 3 });

            // Assert
            var dt = new DataTable();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("C")
            });

            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.DataTables(datatable, dt);
        }

        [Fact]
        public void Load_excel_with_duplicated_columns()
        {
            // Act
            var datatable = ExcelReader.ToDataTable(@"TestData\DuplicatedColumns.xlsx", new() { AllowDuplicateColumns = true });

            // Assert
            var dt = new DataTable();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("B_" + Guid.NewGuid().ToString("N")),
                new DataColumn("C")
            });

            dt.AddRow(new object[] { "1", "2", "3", "4" });
            dt.AddRow(new object[] { "1", "2", "3", "4" });
            dt.AddRow(new object[] { "1", "2", "3", "4" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.ColumnsWithDuplication(datatable, dt, duplicatedColumnNumber: 3);
            DataTableAssert.Rows(datatable, dt);
        }

        [Fact]
        public void Load_excel_without_allowed_duplicated_columns_option_throws_exception()
        {
            // Act
            static void act() =>
                ExcelReader.ToDataTable(@"TestData\DuplicatedColumns.xlsx",
                    new() { AllowDuplicateColumns = false });

            // Assert
            DuplicateNameException exception = Assert.Throws<DuplicateNameException>(act);
        }

        [Fact]
        public void Load_excel_via_sheet_name()
        {
            // Act
            var datatable = ExcelReader.ToDataTable(@"TestData\CustomSheetNameAndPosition.xlsx", new() { SheetName = "Custom Sheet Name" });

            // Assert
            var dt = new DataTable();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("C")
            });

            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.DataTables(datatable, dt);
        }

        [Fact]
        public void Load_excel_via_sheet_position()
        {
            // Act
            var datatable = ExcelReader.ToDataTable(@"TestData\CustomSheetNameAndPosition.xlsx", new() { SheetIndex = 1 });

            // Assert
            var dt = new DataTable();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("C")
            });

            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.DataTables(datatable, dt);
        }

        [Fact]
        public void Load_excel_via_sheet_position_and_sheet_name_throws_exception()
        {
            // Act
            static void act() =>
                ExcelReader.ToDataTable(@"TestData\CustomSheetNameAndPosition.xlsx",
                    new() { SheetName = "Custom Sheet Name", SheetIndex = 1 });

            // Assert
            ArgumentException exception = Assert.Throws<ArgumentException>(act);
        }

        [Fact]
        public void Load_excel_and_throws_an_exception_if_header_row_is_empty()
        {
            // Act
            static void act() =>
                ExcelReader.ToDataTable(@"TestData\EmptyRowsFromTop.xlsx",
                    new() { HeaderRowIndex = 0 });

            // Assert
            ArgumentException exception = Assert.Throws<ArgumentException>(act);
        }

        [Fact]
        public void Load_excel_with_max_column_option_that_is_larger_than_the_existing_columns()
        {
            // Act
            var datatable = ExcelReader.ToDataTable(@"TestData\AdditionalColumns.xlsx", 
                new() { MaxColumns = 20 });

            // Assert
            var dt = new DataTable();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("C"),
                new DataColumn("D"),
                new DataColumn("E")
            });

            dt.AddRow(new object[] { "1", "2", "3", "4", "5" });
            dt.AddRow(new object[] { "1", "2", "3", "4", "5" });
            dt.AddRow(new object[] { "1", "2", "3", "4", "5" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.DataTables(datatable, dt);
        }

        [Fact]
        public void Load_excel_with_known_edge_cases()
        {
            // Act
            // Excel has first sheet hidden.
            var datatable = ExcelReader.ToDataTable(@"TestData\KnownEdgeCases.xlsx", new() { HeaderRowIndex = 1, SheetIndex = 1 });

            // Assert
            var dt = new DataTable();

            var dateValue = DateTime.ParseExact("28/08/2021", "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("Column1"),
                new DataColumn("B"),
                new DataColumn("B_" + Guid.NewGuid().ToString("N")),
                new DataColumn("C"),
                new DataColumn("Column2"),
                new DataColumn("D"),
                new DataColumn("TRUE"),
                new DataColumn("1"),
                new DataColumn(dateValue),
                new DataColumn("12.56"),
            });

            dt.AddRow(new object[] { "1", "Value of B4", "3", "", "5", "", "7", "", "", "1325,48" });
            dt.AddRow(new object[] { "", "", "3", "", "4", "", "", "False", "", "" });
            dt.AddRow(new object[] { "1", "7", "3", "4", dateValue, "6", "7", "", "1234.4895", "" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.ColumnsWithDuplication(datatable, dt, duplicatedColumnNumber: 3);
            DataTableAssert.Rows(datatable, dt);
        }

        [Fact]
        public void Load_xls_excel_with_known_edge_cases()
        {
            // Act
            // Excel has first sheet hidden.
            var datatable = ExcelReader.ToDataTable(@"TestData\KnownEdgeCases.xls",
                new() { HeaderRowIndex = 1, SheetIndex = 1, Format = ExcelReaderFormat.Xls });

            // Assert
            var dt = new DataTable();

            var dateValue = DateTime.ParseExact("28/08/2021", "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("Column1"),
                new DataColumn("B"),
                new DataColumn("B_" + Guid.NewGuid().ToString("N")),
                new DataColumn("C"),
                new DataColumn("Column2"),
                new DataColumn("D"),
                new DataColumn("TRUE"),
                new DataColumn("1"),
                new DataColumn(dateValue),
                new DataColumn("12.56"),
            });

            dt.AddRow(new object[] { "1", "Value of B4", "3", "", "5", "", "7", "", "", "1325,48" });
            dt.AddRow(new object[] { "", "", "3", "", "4", "", "", "False", "", "" });
            dt.AddRow(new object[] { "1", "7", "3", "4", dateValue, "6", "7", "", "1234.4895", "" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.ColumnsWithDuplication(datatable, dt, duplicatedColumnNumber: 3);
            DataTableAssert.Rows(datatable, dt);
        }

        [Fact]
        public void Load_xlsm_excel_with_known_edge_cases()
        {
            // Act
            // Excel has first sheet hidden.
            var datatable = ExcelReader.ToDataTable(@"TestData\KnownEdgeCases.xlsm",
                new() { HeaderRowIndex = 1, SheetIndex = 1, Format = ExcelReaderFormat.Xlsm });

            // Assert
            var dt = new DataTable();

            var dateValue = DateTime.ParseExact("28/08/2021", "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("Column1"),
                new DataColumn("B"),
                new DataColumn("B_" + Guid.NewGuid().ToString("N")),
                new DataColumn("C"),
                new DataColumn("Column2"),
                new DataColumn("D"),
                new DataColumn("TRUE"),
                new DataColumn("1"),
                new DataColumn(dateValue),
                new DataColumn("12.56"),
            });

            dt.AddRow(new object[] { "1", "Value of B4", "3", "", "5", "", "7", "", "", "1325,48" });
            dt.AddRow(new object[] { "", "", "3", "", "4", "", "", "False", "", "" });
            dt.AddRow(new object[] { "1", "7", "3", "4", dateValue, "6", "7", "", "1234.4895", "" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.ColumnsWithDuplication(datatable, dt, duplicatedColumnNumber: 3);
            DataTableAssert.Rows(datatable, dt);
        }
    }
}
