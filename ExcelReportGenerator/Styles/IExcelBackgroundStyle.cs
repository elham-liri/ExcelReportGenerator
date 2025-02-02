using OfficeOpenXml.Style;

namespace ExcelReportGenerator.Styles
{
    /// <summary>
    /// determine Background style
    /// </summary>
    public interface IExcelBackgroundStyle
    {
        /// <summary>
        /// fill pattern
        /// </summary>
        public ExcelFillStyle? BackgroundStyle { get; set; }
        
        /// <summary>
        /// background color can be in one of these forms:
        /// example1 : black
        /// example2 : #000000
        /// example3 : rbg(0,0,0)
        /// example4 : rgba(0,0,0,1)
        /// </summary>
        public string? BackgroundColor { get; set; }

    }
}
