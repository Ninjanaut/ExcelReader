using System;

namespace Ninjanaut.IO
{
    public class ExcelReaderOptions
    {
        public ExcelReaderFormat Format { get; set; }
        public int? SheetIndex { get; set; }
        public string SheetName { get; set; }
        public int? MaxColumns { get; set; } 
        public int? HeaderRowIndex { get; set; }
        public bool RemoveEmptyRows { get; set; }
        public bool AllowDuplicateColumns { get; set; }

        public ExcelReaderOptions()
        {
            Format = ExcelReaderFormat.Xlsx;
            RemoveEmptyRows = true;
            AllowDuplicateColumns = true;
            HeaderRowIndex = 0;
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
