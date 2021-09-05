using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.IO;
using System.Linq;

namespace Ninjanaut.IO
{
    public static class ExcelReader
    {
        /// <summary>
        /// Returns datatable object from the excel file with values retrieved as strings.
        /// </summary>
        /// <param name="bytes">Excel file bytes.</param>
        /// <param name="options">Settings you might want to change.</param>
        public static DataTable ToDataTable(byte[] bytes, ExcelReaderOptions options = null)
        {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));

            options = LoadOptions(options);

            IWorkbook workbook;

            using var stream = new MemoryStream(bytes);

            workbook = LoadWorkbook(stream, options);

            ISheet sheet = LoadSheet(options, workbook);

            return ToDataTable(sheet, options);
        }


        /// <summary>
        /// Returns datatable object from the excel file with values retrieved as strings.
        /// </summary>
        /// <param name="path">Relative or absolute path to the excel file.</param>
        /// <param name="options">Settings you might want to change.</param>
        public static DataTable ToDataTable(string path, ExcelReaderOptions options = null)
        {
            if (string.IsNullOrEmpty(path)) throw new ArgumentNullException(nameof(path));

            options = LoadOptions(options);

            IWorkbook workbook;

            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                workbook = LoadWorkbook(stream, options);
            }

            ISheet sheet = LoadSheet(options, workbook);

            return ToDataTable(sheet, options);
        }

        private static ExcelReaderOptions LoadOptions(ExcelReaderOptions options)
        {
            if (options == null)
            {
                return new ExcelReaderOptions();
            }

            options.Validate();

            return options;
        }

        private static IWorkbook LoadWorkbook(Stream stream, ExcelReaderOptions options)
        {
            return options.Format switch
            {
                ExcelReaderFormat.Xlsx => new XSSFWorkbook(stream),
                ExcelReaderFormat.Xlsm => new XSSFWorkbook(stream),
                ExcelReaderFormat.Xls => new HSSFWorkbook(stream),
                _ => throw new NotSupportedException(),
            };
        }

        private static ISheet LoadSheet(ExcelReaderOptions options, IWorkbook workbook)
        {
            if (!string.IsNullOrEmpty(options.SheetName))
            {
                return workbook.GetSheet(options.SheetName);
            }

            if (options.SheetIndex != null)
            {
                return workbook.GetSheetAt((int)options.SheetIndex);
            }

            return workbook.GetSheetAt(0);
        }

        private static DataTable ToDataTable(ISheet sheet, ExcelReaderOptions options)
        {
            if (sheet is null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            if (sheet.LastRowNum <= options.HeaderRowIndex)
            {
                throw new ArgumentOutOfRangeException(nameof(options), "HeaderRowIndex is not valid.");
            }

            var formulaEvaluator = sheet.Workbook.GetFormulaEvaluator();

            var dataTable = new DataTable(sheet.SheetName);

            foreach (IRow row in sheet)
            {
                if (row is null || row.RowNum < options.HeaderRowIndex) continue;

                if (row.RowNum == options.HeaderRowIndex)
                {
                    LoadHeader(formulaEvaluator, dataTable, row, options);
                }
                else
                {
                    if (dataTable.Columns.Count == 0)
                    {
                        throw new ArgumentException("HeaderRowIndex is not valid.", nameof(options));
                    }
                    LoadRow(formulaEvaluator, dataTable, row, options);
                }
            }

            return dataTable;

            static void LoadHeader(IFormulaEvaluator formulaEvaluator, DataTable dataTable, IRow row, ExcelReaderOptions options)
            {
                for (var columnIndex = 0; columnIndex < row.LastCellNum; columnIndex++)
                {
                    var cell = row.GetCell(columnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);

                    string columnName = cell.GetCellAsString(formulaEvaluator)?.ToString().Trim();

                    try
                    {
                        dataTable.Columns.Add(columnName);
                    }
                    catch (DuplicateNameException)
                    {
                        if (!options.AllowDuplicateColumns) throw;

                        dataTable.Columns.Add(columnName + "_" + Guid.NewGuid().ToString("N"));
                    }

                    if (options.MaxColumns != null && cell.ColumnIndex + 1 == options.MaxColumns) break;
                }
            }

            static void LoadRow(IFormulaEvaluator formulaEvaluator, DataTable dataTable, IRow row, ExcelReaderOptions options)
            {
                var dataRow = dataTable.NewRow();

                for (var columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
                {
                    var cell = row.GetCell(columnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    dataRow[columnIndex] = cell.GetCellAsString(formulaEvaluator);
                }

                if (options.RemoveEmptyRows)
                {
                    var rowContainsData =
                        dataRow.ItemArray.Any(value => value != DBNull.Value && !string.IsNullOrEmpty((string)value));

                    if (rowContainsData)
                    {
                        dataTable.Rows.Add(dataRow);
                    }
                }
                else
                {
                    dataTable.Rows.Add(dataRow);
                }
            }
        }
    }
}
