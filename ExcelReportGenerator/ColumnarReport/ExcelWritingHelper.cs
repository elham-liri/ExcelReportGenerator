using ExcelReportGenerator.Styles;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace ExcelReportGenerator.ColumnarReport
{
    public static class ExcelWritingHelper
    {
        public static ExcelWorksheet CreateSheet(this ExcelPackage package, string sheetName)
        {
            var sheet = package.Workbook.Worksheets.Add(sheetName);
            return sheet;
        }

        public static ExcelWorksheet SetSheetDefaultProperties(this ExcelWorksheet sheet,
            IExcelReportProperties? defaultProperties)
        {
            if (defaultProperties == null) return sheet;

            return sheet.SetSheetDirection(defaultProperties.RightToLeft)
                  .SetSheetFrozenCells(defaultProperties.FrozenColumns, defaultProperties.FrozenRows)
                  .SetSheetProtectionProperties(defaultProperties.IsProtected, defaultProperties.ProtectionPassword,
                      defaultProperties.AllowDeleteRows, defaultProperties.AllowDeleteColumns,
                      defaultProperties.AllowSelectLockedCells);
        }

        public static ExcelWorksheet SetSheetDirection(this ExcelWorksheet sheet, bool rightToLeft)
        {
            sheet.View.RightToLeft = rightToLeft;
            return sheet;
        }

        public static ExcelWorksheet SetSheetFrozenCells(this ExcelWorksheet sheet, int frozenColumns, int frozenRows)
        {
            if (frozenColumns > 0 || frozenRows > 0)
                sheet.View.FreezePanes(frozenRows, frozenColumns);

            return sheet;
        }

        public static ExcelWorksheet SetSheetProtectionProperties(this ExcelWorksheet sheet, bool isProtected,
            string? password, bool allowDeleteRows, bool allowDeleteColumns, bool allowSelectLockedCells)
        {
            if (!isProtected) return sheet;

            sheet.Protection.IsProtected = true;

            if (!string.IsNullOrWhiteSpace(password))
                sheet.Protection.SetPassword(password);

            sheet.Protection.AllowDeleteRows = allowDeleteRows;
            sheet.Protection.AllowDeleteColumns = allowDeleteColumns;
            sheet.Protection.AllowSelectLockedCells = allowSelectLockedCells;

            return sheet;
        }

        public static ExcelWorksheet AddHeaderRows(this ExcelWorksheet sheet, IEnumerable<IExcelReportHeader> headers,
            IEnumerable<IExcelReportColumn> columns)
        {
            var enableHeaders = headers.Where(a => !a.Disabled).OrderBy(a => a.Order).ToList();

            var rowNumber = 1;
            foreach (var header in enableHeaders)
            {
                sheet = header.ColumnBased
                    ? sheet.AddColumnBasedHeader(header, columns, rowNumber)
                    : sheet.AddNonColumnBasedHeader(header, rowNumber);

                rowNumber++;
            }

            return sheet;
        }

        public static ExcelWorksheet AddColumnBasedHeader(this ExcelWorksheet sheet, IExcelReportHeader header,
            IEnumerable<IExcelReportColumn> columns, int rowNumber)
        {
            var verticalColumns = columns.Where(a => !a.IsHorizontal).OrderBy(a => a.Order).ToList();
            var colsCount = verticalColumns.Count;

            using var range = sheet.Cells[rowNumber, 1, rowNumber, colsCount];
            for (var i = 0; i < colsCount; i++)
            {
                range[rowNumber, i + 1].Value = verticalColumns[i].DisplayName;
            }

            return sheet;
        }

        public static ExcelWorksheet AddNonColumnBasedHeader(this ExcelWorksheet sheet, IExcelReportHeader header,
            int rowNumber)
        {
            var columnNumber = 1;
            foreach (var headerPart in header.HeaderParts)
            {
                var range = sheet.Cells[rowNumber, columnNumber, rowNumber, columnNumber + headerPart.Colspan - 1];
                range.Value = headerPart.Value;

                columnNumber += headerPart.Colspan;
            }

            return sheet;
        }


        public static ExcelWorksheet SetDefaultSheetStyle(this ExcelWorksheet sheet, IExcelReportProfile reportProfile, int dataSetCount)
        {
            var rows = reportProfile.GetFinalRowsCount(dataSetCount);
            var cols = reportProfile.GetVerticalColumnsCount();

            sheet
                .SetDefaultHeightForRows(rows, reportProfile.DefaultStyle.DefaultRowHeight)
                .SetDefaultWidthForColumns(cols, reportProfile.DefaultStyle.DefaultColumnWidth)
                .ApplyRowColorPallet()
                .ApplyColumnColorPallet()
                .Cells[1, rows, 1, cols]
                .SetFontStyle(reportProfile.DefaultStyle.FontStyle)
                .SetBackgroundStyle(reportProfile.DefaultStyle.BackgroundStyle)
                .SetBorderStyle(reportProfile.DefaultStyle.BorderStyle);



            return sheet;
        }

        public static ExcelRange SetFontStyle(this ExcelRange range, IExcelFontStyle? fontStyle)
        {
            if (fontStyle == null) return range;

            if (!string.IsNullOrWhiteSpace(fontStyle.Name))
                range.Style.Font.Name = fontStyle.Name;

            if (fontStyle.Size.HasValue && fontStyle.Size > 0)
                range.Style.Font.Size = fontStyle.Size.Value;

            range.Style.Font.Bold = fontStyle.Bold;
            range.Style.Font.Italic = fontStyle.Italic;
            range.Style.Font.UnderLine = fontStyle.UnderLine;

            if (!string.IsNullOrWhiteSpace(fontStyle.Color))
                range.Style.Font.Color.SetColor(GetColor(fontStyle.Color));

            if (fontStyle.HorizontalAlignment.HasValue)
                range.Style.HorizontalAlignment = fontStyle.HorizontalAlignment.Value;

            if (fontStyle.VerticalAlignment.HasValue)
                range.Style.VerticalAlignment = fontStyle.VerticalAlignment.Value;

            return range;
        }

        public static ExcelRange SetBackgroundStyle(this ExcelRange range, IExcelBackgroundStyle? backgroundStyle)
        {
            if (backgroundStyle == null) return range;

            if (backgroundStyle.Style.HasValue)
                range.Style.Fill.PatternType = backgroundStyle.Style.Value;

            if (!string.IsNullOrWhiteSpace(backgroundStyle.Color))
                range.Style.Fill.BackgroundColor.SetColor(GetColor(backgroundStyle.Color));

            return range;
        }

        public static ExcelRange SetBorderStyle(this ExcelRange range, IExcelBorderStyle? borderStyle)
        {
            if (borderStyle == null) return range;

            if (borderStyle.BorderAround)
            {
                if (borderStyle.AroundStyle == null) return range;

                if (string.IsNullOrWhiteSpace(borderStyle.AroundColor))
                    range.Style.Border.BorderAround(borderStyle.AroundStyle.Value);
                else
                    range.Style.Border.BorderAround(borderStyle.AroundStyle.Value, GetColor(borderStyle.AroundColor));

                return range;
            }

            if (borderStyle.TopStyle.HasValue)
            {
                range.Style.Border.Top.Style = borderStyle.TopStyle.Value;

                if (!string.IsNullOrWhiteSpace(borderStyle.TopColor))
                    range.Style.Border.Top.Color.SetColor(GetColor(borderStyle.TopColor));
            }

            if (borderStyle.LeftStyle.HasValue)
            {
                range.Style.Border.Left.Style = borderStyle.LeftStyle.Value;

                if (!string.IsNullOrWhiteSpace(borderStyle.LeftColor))
                    range.Style.Border.Left.Color.SetColor(GetColor(borderStyle.LeftColor));
            }

            if (borderStyle.BottomStyle.HasValue)
            {
                range.Style.Border.Bottom.Style = borderStyle.BottomStyle.Value;

                if (!string.IsNullOrWhiteSpace(borderStyle.BottomColor))
                    range.Style.Border.Bottom.Color.SetColor(GetColor(borderStyle.BottomColor));
            }

            if (borderStyle.RightStyle.HasValue)
            {
                range.Style.Border.Right.Style = borderStyle.RightStyle.Value;

                if (!string.IsNullOrWhiteSpace(borderStyle.RightColor))
                    range.Style.Border.Right.Color.SetColor(GetColor(borderStyle.RightColor));
            }

            return range;
        }

        public static ExcelWorksheet SetDefaultHeightForRows(this ExcelWorksheet sheet, int rows, double? defaultHeight)
        {
            if (defaultHeight is not > 0) return sheet;

            for (var i = 0; i < rows; i++)
            {
                sheet.Row(i + 1).Height = defaultHeight.Value;
            }

            return sheet;
        }

        public static ExcelWorksheet SetDefaultWidthForColumns(this ExcelWorksheet sheet, int cols, double? defaultWidth)
        {
            if (defaultWidth is not > 0) return sheet;

            for (int i = 0; i < cols; i++)
            {
                sheet.Column(i + 1).Width = defaultWidth.Value;
            }

            return sheet; }

        public static ExcelWorksheet ApplyRowColorPallet(this ExcelWorksheet sheet)
        {
            //TODO
            return sheet;
        }

        public static ExcelWorksheet ApplyColumnColorPallet(this ExcelWorksheet sheet)
        {
            //TODO
            return sheet;
        }

        public static Color GetColor(this string code)
        {
            if (code.Contains("rgba"))
            {
                code = code.Replace("rgba", string.Empty).Replace("(", string.Empty).Replace(")", string.Empty);
                var rgba = code.Split(',');
                rgba[3] = rgba[3].Contains("0.") ? rgba[3].Replace("0.", string.Empty) : rgba[3];
                return Color.FromArgb(int.Parse(rgba[3]), int.Parse(rgba[0]), int.Parse(rgba[1]), int.Parse(rgba[2]));
            }

            if (code.Contains("rgb"))
            {
                code = code.Replace("rgb", string.Empty).Replace("(", string.Empty).Replace(")", string.Empty);
                var rgb = code.Split(',');
                return Color.FromArgb(int.Parse(rgb[0]), int.Parse(rgb[1]), int.Parse(rgb[2]));
            }
            return System.Drawing.ColorTranslator.FromHtml(code);
        }




        public static int GetVerticalColumnsCount(this IExcelReportProfile profile)
        {
            return profile.Columns.Count(a => !a.IsHorizontal);
        }

        public static int GetFinalRowsCount(this IExcelReportProfile profile, int dataSetCount)
        {
            var finalRow = dataSetCount + profile.Headers.Count(a => !a.Disabled);
            finalRow += dataSetCount * profile.Columns.Count(a => a.IsHorizontal);

            if (profile.Columns.Any(a => a.ShowTotalSum)) finalRow += 1;

            return finalRow;
        }

    }
}
