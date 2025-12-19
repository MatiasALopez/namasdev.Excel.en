using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;

using namasdev.Core.Types;
using namasdev.Core.Validation;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace namasdev.Excel
{
    public abstract class ExcelRecord
    {
        private List<string> _errors;

        public ExcelRecord(ExcelWorksheet excelWorksheet, int row,
            bool includeWorksheetNameInError = true)
        {
            Validator.ValidateRequiredArgumentAndThrow(excelWorksheet, nameof(excelWorksheet));

            ExcelWorksheet = excelWorksheet;
            Row = row;

            _errors = new List<string>();

            IncludeWorksheetNameInError = includeWorksheetNameInError;
        }

        public ExcelWorksheet ExcelWorksheet { get; private set; }
        public int Row { get; private set; }
        public IEnumerable<string> Errors
        {
            get { return _errors.AsReadOnly(); }
        }

        public bool IncludeWorksheetNameInError { get; set; }

        public virtual bool IsEmpty { get; }

        public bool IsValid
        {
            get { return _errors.Count == 0; }
        }

        protected string GetString(int column, string fieldDescription,
            bool required = true,
            int? maxLength = null, int? exactLength = null,
            bool trimValue = false)
        {
            var cell = GetCell(column);
            var value = cell.Text.ValueNotEmptyOrNull();

            if (!Validator.ValidateString(value, fieldDescription, required,
                out string errorMessage,
                maxLength: maxLength, exactLength: exactLength))
            {
                AddError(cell, errorMessage);
            }
            else if (trimValue)
            {
                value = value.Trim();
            }

            return value;
        }

        protected int? GetInt(int column, string fieldDescription,
            bool required = true)
        {
            var cell = GetCell(column);
            var value = cell.Value;

            if (value != null
                && !String.IsNullOrWhiteSpace(Convert.ToString(value)))
            {
                try
                {
                    return Convert.ToInt32(value);
                }
                catch (Exception)
                {
                    AddError(cell, Validator.Messages.IntegerInvalid(fieldDescription));
                    return null;
                }
            }
            else
            {
                if (required)
                {
                    AddError(cell, Validator.Messages.Required(fieldDescription));
                }

                return null;
            }
        }

        protected string GetStringYValidarFormatoCorreo(int column, string fieldDescription, 
            bool required = true)
        {
            var cell = GetCell(column);
            string stringValue = GetString(column, fieldDescription, required);
            if (!string.IsNullOrEmpty(stringValue))
            {
                string errorMessage;
                if (!Validator.ValidateEmail(stringValue, fieldDescription, 
                    required: required,
                    out errorMessage))
                {
                    AddError(cell, errorMessage);
                }
            }

            return stringValue;
        }

        protected short? GetShort(int column, string fieldDescription, 
            bool required = true)
        {
            var cell = GetCell(column);
            var value = cell.Value;

            if (value != null
                && !String.IsNullOrWhiteSpace(Convert.ToString(value)))
            {
                try
                {
                    return Convert.ToInt16(value);
                }
                catch (Exception)
                {
                    AddError(cell, Validator.Messages.ShortInvalid(fieldDescription));
                    return null;
                }
            }
            else
            {
                if (required)
                {
                    AddError(cell, Validator.Messages.Required(fieldDescription));
                }

                return null;
            }
        }

        protected long? GetLong(int column, string fieldDescription, 
            bool required = true)
        {
            var cell = GetCell(column);
            var value = cell.Value;

            if (value != null
                && !String.IsNullOrWhiteSpace(Convert.ToString(value)))
            {
                try
                {
                    return Convert.ToInt64(value);
                }
                catch (Exception)
                {
                    AddError(cell, Validator.Messages.LongInvalid(fieldDescription));
                    return null;
                }
            }
            else
            {
                if (required)
                {
                    AddError(cell, Validator.Messages.Required(fieldDescription));
                }

                return null;
            }
        }

        protected double? GetDouble(int column, string fieldDescription, 
            bool required = true)
        {
            var cell = GetCell(column);
            var value = cell.Value;

            if (value != null
                && !String.IsNullOrWhiteSpace(Convert.ToString(value)))
            {
                try
                {
                    return Convert.ToDouble(value);
                }
                catch (Exception)
                {
                    AddError(cell, Validator.Messages.NumberInvalid(fieldDescription));
                    return null;
                }
            }
            else
            {
                if (required)
                {
                    AddError(cell, Validator.Messages.Required(fieldDescription));
                }

                return null;
            }
        }

        protected decimal? GetDecimal(int column, string fieldDescription,
            bool required = true)
        {
            var cell = GetCell(column);
            var value = cell.Value;

            if (value != null
                && !String.IsNullOrWhiteSpace(Convert.ToString(value)))
            {
                try
                {
                    return Convert.ToDecimal(value);
                }
                catch (Exception)
                {
                    AddError(cell, Validator.Messages.NumberInvalid(fieldDescription));
                    return null;
                }
            }
            else
            {
                if (required)
                {
                    AddError(cell, Validator.Messages.Required(fieldDescription));
                }

                return null;
            }
        }

        protected DateTime? GetDateTime(int column, string fieldDescription, 
            bool required = true)
        {
            var cell = GetCell(column);
            var value = cell.Value;

            if (value != null
                && !String.IsNullOrWhiteSpace(Convert.ToString(value).Replace("-", "")))
            {
                try
                {
                    return Convert.ToDateTime(value);
                }
                catch (Exception)
                {
                    try
                    {
                        double dato = Convert.ToDouble(value);
                        return DateTime.FromOADate(dato);
                    }
                    catch (Exception)
                    {
                        AddError(cell, Validator.Messages.DateTimeInvalid(fieldDescription));
                        return null;
                    }
                }
            }
            else
            {
                if (required)
                {
                    AddError(cell, Validator.Messages.Required(fieldDescription));
                }

                return null;
            }
        }

        protected TimeSpan? GetTimeSpan(int column, string fieldDescription, 
            bool required = true)
        {
            var cell = GetCell(column);
            var value = cell.Value;

            if (value != null
                && !String.IsNullOrWhiteSpace(Convert.ToString(value).Replace("-", "")))
            {
                try
                {
                    return TimeSpan.Parse(Convert.ToString(value));
                }
                catch (Exception)
                {
                    try
                    {
                        return Convert.ToDateTime(value).TimeOfDay;
                    }
                    catch (Exception)
                    {
                        AddError(cell, Validator.Messages.TimeInvalid(fieldDescription));
                        return null;
                    }
                }
            }
            else
            {
                if (required)
                {
                    AddError(cell, Validator.Messages.Required(fieldDescription));
                }

                return null;
            }
        }

        protected bool? GetBoolean(int column, string fieldDescription, 
            bool required = true)
        {
            var cell = GetCell(column);
            var value = cell.Value;

            if (value != null)
            {
                bool? boolValue = value as bool?;
                if (boolValue.HasValue)
                {
                    return boolValue.Value;
                }

                var stringValue = value.ToString();
                return string.Equals(stringValue, Formatter.YES, StringComparison.CurrentCultureIgnoreCase);
            }
            else
            {
                if (required)
                {
                    AddError(cell, Validator.Messages.Required(fieldDescription));
                }

                return null;
            }
        }

        protected short? GetMonthNumber(int column, string fieldDescription, 
            bool required = true)
        {
            var value = GetString(column, fieldDescription, required: required);
            if (String.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            try
            {
                short month;
                if (short.TryParse(value, out month))
                {
                    return month;
                }

                var monthNames = new List<string>(CultureInfo.CurrentCulture.DateTimeFormat.MonthNames);
                short monthNumber = (short)(monthNames.IndexOf(value.ToLower()) + 1);
                if (monthNumber < 1 || monthNumber > 12)
                {
                    throw new Exception(Validator.Messages.EntityNotFound("Month", monthNumber));
                }

                return monthNumber;
            }
            catch (Exception)
            {
                AddError(GetCell(column), Validator.Messages.TypeInvalid(fieldDescription, "month"));
                return null;
            }
        }

        protected ExcelRange GetCell(int column)
        {
            return ExcelWorksheet.Cells[Row, column];
        }

        protected void AddError(ExcelRange cell, string error)
        {
            if (cell != null)
            {
                _errors.Add($"[{(IncludeWorksheetNameInError ? cell.FullAddress : cell.Address)}] {error}");
            }
            else
            {
                _errors.Add(error);
            }
        }

        protected void SetCellValue(int column, object value,
            bool? wrapText = null,
            string numberFormat = null)
        {
            ExcelWorksheet.Cells[Row, column].Value = value;

            if (wrapText.HasValue)
            {
                ExcelWorksheet.Cells[Row, column].Style.WrapText = wrapText.Value;
            }

            if (!string.IsNullOrWhiteSpace(numberFormat))
            {
                ExcelWorksheet.Cells[Row, column].Style.Numberformat.Format = numberFormat;
            }
        }

        public void ApplyStyle(int column,
            ExcelHorizontalAlignment? horizontalAlignment = null, ExcelVerticalAlignment? verticalAlignment = null,
            bool? fontBold = null, bool? autoFit = null,
            Color? textColor = null, Color? backgroundColor = null, Color? borderColor = null)
        {
            ExcelWorksheet.Cells[Row, column].ApplyStyle(
                horizontalAlignment: horizontalAlignment,
                verticalAlignment: verticalAlignment,
                fontBold: fontBold,
                autoFit: autoFit,
                textColor: textColor,
                backgroundColor: backgroundColor,
                borderColor: borderColor);
        }

        public void ApplyStyle(int columnFrom, int columnTo,
            ExcelHorizontalAlignment? horizontalAlignment = null, ExcelVerticalAlignment? verticalAlignment = null,
            bool? fontBold = null, bool? autoFit = null,
            Color? textColor = null, Color? backgroundColor = null, Color? borderColor = null)
        {
            ExcelWorksheet.Cells[Row, columnFrom, Row, columnTo].ApplyStyle(
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
