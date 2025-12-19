using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

using namasdev.Core.Exceptions;
using namasdev.Core.Types;
using namasdev.Core.Validation;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace namasdev.Excel
{
    public class Worksheet
    {
        public Worksheet(ExcelWorkbook workbook, string name)
        {
            Validator.ValidateRequiredArgumentAndThrow(workbook, nameof(workbook));
            Validator.ValidateRequiredArgumentAndThrow(name, nameof(name));

            ExcelWorkbook = workbook;
            Name = name;

            ExcelWorksheet = ExcelWorkbook.Worksheets[name];
            if (ExcelWorksheet == null)
            {
                throw new ExceptionFriendlyMessage(Validator.Messages.EntityNotFound("Worksheet", name));
            }
        }

        public Worksheet(ExcelWorkbook workbook,
            int number = 1)
        {
            Validator.ValidateRequiredArgumentAndThrow(workbook, nameof(workbook));

            ExcelWorkbook = workbook;

            ExcelWorksheet = ExcelWorkbook.Worksheets[number];
            if (ExcelWorksheet == null)
            {
                throw new ExceptionFriendlyMessage(Validator.Messages.EntityNotFound("Worksheet", number));
            }

            Name = ExcelWorksheet.Name;
        }

        public Worksheet(ExcelWorksheet excelWorksheet)
        {
            Validator.ValidateRequiredArgumentAndThrow(excelWorksheet, nameof(excelWorksheet));

            ExcelWorksheet = excelWorksheet;
            ExcelWorkbook = excelWorksheet.Workbook;
            Name = excelWorksheet.Name;
        }

        public ExcelWorkbook ExcelWorkbook { get; private set; }
        public ExcelWorksheet ExcelWorksheet { get; private set; }
        public string Name { get; private set; }

        public void ValidateHeaders(string[] headers,
            int column = 1,
            int row = 1)
        {
            var headersNotFound = new List<string>();

            int headersCount = headers.Length;
            string header;
            ExcelRange cell;
            for (int i = 0; i < headersCount; i++)
            {
                header = headers[i];
                cell = ExcelWorksheet.Cells[row, i + column];

                if (!String.Equals(cell.Text.Trim(), header.Trim(), StringComparison.CurrentCultureIgnoreCase))
                {
                    headersNotFound.Add($"{header} ({cell.Address})");
                }
            }

            if (headersNotFound.Any())
            {
                throw new ExceptionFriendlyMessage($"[{Name}] Headers not found: {Formatter.List(headersNotFound, ", ")}.");
            }
        }

        public void ApplyStyle(int row, int column,
            ExcelHorizontalAlignment? horizontalAlignment = null, ExcelVerticalAlignment? verticalAlignment = null,
            bool? fontBold = null, bool? autoFit = null, 
            Color? textColor = null, Color? backgroundColor = null, Color? borderColor = null)
        {
            ExcelWorksheet.Cells[row, column].ApplyStyle(
                horizontalAlignment: horizontalAlignment,
                verticalAlignment: verticalAlignment,
                fontBold: fontBold,
                autoFit: autoFit,
                textColor: textColor,
                backgroundColor: backgroundColor,
                borderColor: borderColor);
        }

        public void ApplyStyle(int rowFrom, int columnFrom, int rowTo, int columnTo,
            ExcelHorizontalAlignment? horizontalAlignment = null, ExcelVerticalAlignment? verticalAlignment = null,
            bool? fontBold = null, bool? autoFit = null, 
            Color? textColor = null, Color? backgroundColor = null, Color? borderColor = null)
        {
            ExcelWorksheet.Cells[rowFrom, columnFrom, rowTo, columnTo].ApplyStyle(
                horizontalAlignment: horizontalAlignment,
                verticalAlignment: verticalAlignment,
                fontBold: fontBold,
                autoFit: autoFit,
                textColor: textColor,
                backgroundColor: backgroundColor,
                borderColor: borderColor);
        }

        public void ApplyStyle(string cellRange,
            ExcelHorizontalAlignment? horizontalAlignment = null, ExcelVerticalAlignment? verticalAlignment = null,
            bool? fontBold = null, bool? autoFit = null, 
            Color? textColor = null, Color? backgroundColor = null, Color? borderColor = null)
        {
            ExcelWorksheet.Cells[cellRange].ApplyStyle(
                horizontalAlignment: horizontalAlignment,
                verticalAlignment: verticalAlignment,
                fontBold: fontBold,
                autoFit: autoFit,
                textColor: textColor,
                backgroundColor: backgroundColor,
                borderColor: borderColor);
        }
    }
}
