using ExcelReportGenerator.Styles;

namespace ExcelReportGenerator.ColumnarReport
{
    /// <summary>
    /// General class for excel styles
    /// </summary>
    public interface IExcelReportStyle
    {
        /// <summary>
        /// font style including name,size,color, alignment,...
        /// </summary>
        public IExcelFontStyle? FontStyle { get; set; }

        /// <summary>
        /// border style including color and style for every side
        /// </summary>
        public IExcelBorderStyle? BorderStyle { get; set; }

        /// <summary>
        /// background style including pattern and color
        /// </summary>
        public IExcelBackgroundStyle? BackgroundStyle { get; set; }

    }


}
