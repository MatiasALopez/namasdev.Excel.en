using System.Drawing;

using namasdev.Core.Validation;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace namasdev.Excel
{
    public static class ExcelRangeExtensions
    {
        public static void ApplyStyle(this ExcelRange range,
            ExcelHorizontalAlignment? horizontalAlignment = null, ExcelVerticalAlignment? verticalAlignment = null,
            bool? fontBold = null, bool? autoFit = null, Color? textColor = null, Color? backgroundColor = null, Color? borderColor = null)
        {
            Validator.ValidateRequiredArgumentAndThrow(range, nameof(range));

            if (textColor.HasValue)
            {
                range.Style.Font.Color.SetColor(textColor.Value);
            }

            if (backgroundColor.HasValue)
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(backgroundColor.Value);
            }

            if (borderColor.HasValue)
            {
                var border = range.Style.Border;
                border.Top.Style = border.Bottom.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
                border.Top.Color.SetColor(borderColor.Value);
                border.Bottom.Color.SetColor(borderColor.Value);
                border.Left.Color.SetColor(borderColor.Value);
                border.Right.Color.SetColor(borderColor.Value);
            }

            if (fontBold.HasValue)
            {
                range.Style.Font.Bold = fontBold.Value;
            }

            if (horizontalAlignment.HasValue)
            {
                range.Style.HorizontalAlignment = horizontalAlignment.Value;
            }

            if (verticalAlignment.HasValue)
            {
                range.Style.VerticalAlignment = verticalAlignment.Value;
            }

            if (autoFit == true)
            {
                range.AutoFitColumns();
                range.Style.WrapText = true;
            }
        }
    }
}
