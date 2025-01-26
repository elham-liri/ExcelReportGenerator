using OfficeOpenXml;
using ExcelReportGenerator.Interfaces;
using OfficeOpenXml.Style;
using System.Drawing;

namespace ExcelReportGenerator
{
    public static class ExcelWritingHelper
    {
        internal static string FirstCharToUpper(this string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;
            return input.First().ToString().ToUpper() + string.Join("", input.Skip(1));
        }

        public static ExcelWorksheet CreateSheet(this ExcelPackage package, string sheetName)
        {
            sheetName = sheetName.Contains("-") ? sheetName.Remove(sheetName.LastIndexOf('-')) : sheetName;
            var sheet = package.Workbook.Worksheets.Add(sheetName);
            return sheet;
        }

        public static ExcelWorksheet SetSheetDefaults(this ExcelWorksheet sheet,
            IExcelProfileDefaultProperties? defaultProperties)
        {
            if (defaultProperties == null) return sheet;

            sheet.Cells.Style.Font.Name = defaultProperties.DefaultFontName;
            sheet.Cells.Style.Font.Size = defaultProperties.DefaultFontSize;
            sheet.View.RightToLeft = defaultProperties.RightToLeft;
            sheet.DefaultRowHeight = defaultProperties.DefaultRowHeight;

            if (defaultProperties.FrozenColumns > 0 || defaultProperties.FrozenRows > 0)
                sheet.View.FreezePanes(defaultProperties.FrozenRows, defaultProperties.FrozenColumns);

            if (!defaultProperties.IsProtected) return sheet;

            sheet.Protection.IsProtected = defaultProperties.IsProtected;
            sheet.Protection.SetPassword(Guid.NewGuid().ToString());

            sheet.Protection.AllowDeleteRows = defaultProperties.AllowDeleteRows;
            sheet.Protection.AllowDeleteColumns = defaultProperties.AllowDeleteColumns;
            sheet.Protection.AllowSelectLockedCells = defaultProperties.AllowSelectLockedCells;

            return sheet;
        }

        public static ExcelWorksheet AddHeaderRow(this ExcelWorksheet sheet, IBaseExcelProfile profile, int rowNumber = 1)
        {
            var headerStyle = profile.HeaderStyle;
            if ( headerStyle==null || !headerStyle.ShowRow) return sheet;

            var columns = profile.GetDisplayedVerticalColumns();
            var colsCount = columns.Count;
            using (var range = sheet.Cells[rowNumber, 1, rowNumber, colsCount])
            {
                for (var i = 0; i < colsCount; i++)
                {
                    range[rowNumber, i + 1].Value = columns[i].NameToShow ?? columns[i].Name;
                    range[rowNumber, i + 1]
                        .SetBorderForCell(headerStyle.BorderStyle, headerStyle.BorderColor);
                }
            }

            sheet.Row(rowNumber).Height = headerStyle.Height ?? profile.DefaultProperties!.DefaultRowHeight;
            sheet.Cells[rowNumber, 1, rowNumber, colsCount].Style.Font.Bold = headerStyle.FontBold;
            sheet.Cells[rowNumber, 1, rowNumber, colsCount].Style.HorizontalAlignment = headerStyle.Alignment;
            sheet.Cells[rowNumber, 1, rowNumber, colsCount]
                .SetBackgroundForCell(headerStyle.FillStyle, headerStyle.BackgroundColor);


            return sheet;
        }

        public static void SetBorderForCell(this ExcelRange range, ExcelBorderStyle borderStyle, Color borderColor)
        {
            if (borderStyle == ExcelBorderStyle.None) return;
            range.Style.Border.BorderAround(borderStyle, borderColor);
        }

        public static void SetBackgroundForCell(this ExcelRange range, ExcelFillStyle fillStyle, Color backgroundColor)
        {
            if (fillStyle == ExcelFillStyle.None) return;

            range.Style.Fill.PatternType = fillStyle;
            range.Style.Fill.BackgroundColor.SetColor(backgroundColor);
        }
    }
}
