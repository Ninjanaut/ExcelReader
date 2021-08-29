using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;

namespace Ninjanaut.IO
{
    public static class NpoiExtensions
    {
        public static IFormulaEvaluator GetFormulaEvaluator(this IWorkbook workbook)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            return workbook switch
            {
                HSSFWorkbook => new HSSFFormulaEvaluator(workbook),
                XSSFWorkbook => new XSSFFormulaEvaluator(workbook),
                _ => throw new NotSupportedException()
            };
        }

        public static object GetCellAsString(this ICell cell, IFormulaEvaluator formulaEvaluator = null)
        {
            if (cell is null || cell.CellType == CellType.Blank || cell.CellType == CellType.Error)
            {
                return null;
            }

            if (cell.CellType == CellType.Numeric)
            {
                if (DateUtil.IsCellDateFormatted(cell))
                {
                    return cell.DateCellValue.ToString();
                }

                return cell.NumericCellValue.ToString();
            }

            if (cell.CellType == CellType.String) return cell.StringCellValue;

            if (cell.CellType == CellType.Boolean) return cell.BooleanCellValue;

            if (cell.CellType == CellType.Formula)
            {
                var evaluatedCellValue = formulaEvaluator?.Evaluate(cell);

                if (evaluatedCellValue != null)
                {
                    if (evaluatedCellValue.CellType == CellType.Blank || evaluatedCellValue.CellType == CellType.Error)
                    {
                        return null;
                    }
                    if (evaluatedCellValue.CellType == CellType.Numeric)
                    {
                        if (DateUtil.IsCellDateFormatted(cell))
                        {
                            return cell.DateCellValue.ToString();
                        }

                        return evaluatedCellValue.NumberValue.ToString();
                    }
                    if (evaluatedCellValue.CellType == CellType.Boolean)
                    {
                        return evaluatedCellValue.BooleanValue.ToString();
                    }
                    if (evaluatedCellValue.CellType == CellType.String)
                    {
                        return evaluatedCellValue.StringValue;
                    }
                    return evaluatedCellValue.FormatAsString();
                }

                return cell.ToString();
            }

            return cell.ToString();
        }
    }
}
