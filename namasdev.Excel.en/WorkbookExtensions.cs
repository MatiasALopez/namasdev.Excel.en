using System;
using System.Collections.Generic;

using OfficeOpenXml;

namespace namasdev.Excel
{
    public static class WorkbookExtensions
    {
        public static void SetNamedRange<T>(this ExcelWorkbook workbook, string worksheetName, string rangeName, int column, IEnumerable<T> values,
            Func<T, object> valueMapper = null, 
            int rowFrom = 1, int rowTo = 9999)
        {
            var worksheet = workbook.Worksheets[worksheetName];
            worksheet.Cells[rowFrom, column, rowTo, column].Value = null;

            int row = rowFrom - 1;
            foreach (var value in values)
            {
                row++;

                if (valueMapper != null)
                {
                    worksheet.Cells[row, column].Value = valueMapper(value);
                }
                else
                {
                    worksheet.Cells[row, column].Value = value;
                }
            }
            workbook.Names[rangeName].Address = worksheet.Cells[rowFrom, column, row, column].FullAddress;
        }
    }
}
