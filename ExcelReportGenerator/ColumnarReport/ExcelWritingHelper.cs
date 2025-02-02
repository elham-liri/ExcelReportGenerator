using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportGenerator.ColumnarReport
{
    public static class ExcelWritingHelper
    {
        public static ExcelWorksheet CreateSheet(this ExcelPackage package, string sheetName)
        {
            sheetName = sheetName.Contains("-") ? sheetName.Remove(sheetName.LastIndexOf('-')) : sheetName;
            var sheet = package.Workbook.Worksheets.Add(sheetName);
            return sheet;
        }

        public static ExcelWorksheet SetSheetDefaultProperties(this ExcelWorksheet sheet,
            IExcelReportProperties? defaultProperties)
        {
            if (defaultProperties == null) return sheet;

            sheet.View.RightToLeft = defaultProperties.RightToLeft;

            if (defaultProperties.FrozenColumns > 0 || defaultProperties.FrozenRows > 0)
                sheet.View.FreezePanes(defaultProperties.FrozenRows, defaultProperties.FrozenColumns);

            if (!defaultProperties.IsProtected) return sheet;

            sheet.Protection.IsProtected = defaultProperties.IsProtected;

            if (defaultProperties.IsProtected && !string.IsNullOrWhiteSpace(defaultProperties.ProtectionPassword))
                sheet.Protection.SetPassword(defaultProperties.ProtectionPassword);

            sheet.Protection.AllowDeleteRows = defaultProperties.AllowDeleteRows;
            sheet.Protection.AllowDeleteColumns = defaultProperties.AllowDeleteColumns;
            sheet.Protection.AllowSelectLockedCells = defaultProperties.AllowSelectLockedCells;

            //TODO: AUTO-FIT SHOULD BE SET AFTER CALCULATIONS
            return sheet;
        }


    }
}
