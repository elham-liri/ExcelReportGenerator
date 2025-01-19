using OfficeOpenXml.Style;
using System.Drawing;

namespace ExcelReportGenerator.Interfaces
{
    public interface IBaseExcelProfileRowStyle
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
