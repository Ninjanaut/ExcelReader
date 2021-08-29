using System.Data;

namespace ExcelReaderTests.Utilities
{
    public static class DataTableExtensions
    {
        public static void AddRow(this DataTable datatable, object[] rowData)
        {
            var row = datatable.NewRow();
            row.ItemArray = rowData;
            datatable.Rows.Add(row);
        }
    }
}
