using OfficeOpenXml;

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
                var range = sheet.Cells[rowNumber, columnNumber, rowNumber, columnNumber + headerPart.Colspan -1];
                range.Value = headerPart.Value;

                columnNumber += headerPart.Colspan;
            }

            return sheet;
        }

    }
}
