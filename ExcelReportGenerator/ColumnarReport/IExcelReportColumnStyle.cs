using ExcelReportGenerator.Styles;

namespace ExcelReportGenerator.ColumnarReport
{
    /// <summary>
    /// Determine styles for a column - these styles are superior to profile style  and will overwrite them for a specific column
    /// </summary>
    public interface IExcelReportColumnStyle : IExcelStyle
    {
        /// <summary>
        /// column width - if autofit is enabled for the column this will be used as minimum width 
        /// </summary>
        public double? Width { get; set; }
    }
}
