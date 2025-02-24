using ExcelReportGenerator.Styles;

namespace ExcelReportGenerator.ColumnarReport
{
    public interface IExcelReportHeaderStyle :IExcelStyle
    {
        public double? Height { get; set; }
    }
}
