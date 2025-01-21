using System.Drawing;
using ExcelReportGenerator.Interfaces;
using OfficeOpenXml.Style;

namespace ExcelReportGenerator.ExcelEntities
{
    public class ExcelProfileCellStyle:IExcelProfileCellStyle
    {
        public ExcelBorderStyle RegularCellBorderStyle { get; set; }
        public Color RegularCellBorderColor { get; set; }
        public ExcelFillStyle HorizontalCellFillStyle { get; set; }
        public Color HorizontalCellBackgroundColor { get; set; }
        public ExcelBorderStyle HorizontalCellBorderStyle { get; set; }
        public Color HorizontalCellBorderColor { get; set; }
    }
}
