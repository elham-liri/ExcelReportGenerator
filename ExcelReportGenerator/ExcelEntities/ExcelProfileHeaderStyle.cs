using System.Drawing;
using ExcelReportGenerator.Interfaces;
using OfficeOpenXml.Style;

namespace ExcelReportGenerator.ExcelEntities
{
    public class ExcelProfileHeaderStyle:IExcelProfileHeaderStyle
    {
        public bool ShowRow { get; set; }
        public double? Height { get; set; }
        public bool FontBold { get; set; }
        public ExcelHorizontalAlignment Alignment { get; set; }
        public ExcelFillStyle FillStyle { get; set; }
        public Color BackgroundColor { get; set; }
        public ExcelBorderStyle BorderStyle { get; set; }
        public Color BorderColor { get; set; }
    }
}
