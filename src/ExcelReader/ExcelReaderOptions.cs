using System;

namespace Ninjanaut.IO
{
    public class ExcelReaderOptions
    {
        /// <summary>
        /// Default format is XLSX.
        /// </summary>
        public ExcelReaderFormat Format { get; set; }

        /// <summary>
        /// The index of the excel sheet that you want to convert to a datatable object.
        /// Setting SheetIndex with SheetName will throw an exception.
        /// If SheetName is not set, the default value is 0.
        /// </summary>
        public int? SheetIndex { get; set; }

        /// <summary>
        /// The name of the excel sheet that you want to convert to a datatable object.
        /// Setting SheetName with SheetIndex will throw an exception.
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// Default value is null. I recommend setting this value so that you don't accidentally load empty columns.
        /// </summary>
        public int? MaxColumns { get; set; }

        /// <summary>
        /// Default value is 0.
        /// </summary>
        public int? HeaderRowIndex { get; set; }

        /// <summary>
        /// Default value is true.
        /// </summary>
        public bool RemoveEmptyRows { get; set; }

        /// <summary>
        /// Default value is true.
        /// </summary>
        public bool AllowDuplicateColumns { get; set; }

        public ExcelReaderOptions()
        {
            Format = ExcelReaderFormat.Xlsx;
            RemoveEmptyRows = true;
            AllowDuplicateColumns = true;
            HeaderRowIndex = 0;
            MaxColumns = null;
        }

        public void Validate()
        {
            if (!string.IsNullOrEmpty(SheetName) && SheetIndex != null)
            {
                throw new ArgumentException("SheetName and SheetIndex cannot be defined at once, please choose one or the other.");
            }

            if (!Enum.IsDefined(typeof(ExcelReaderFormat), Format))
            {
                throw new ArgumentException("Format is not valid ExcelReaderFormat type.");
            }
        }
    }
}
