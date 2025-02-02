using ExcelReportGenerator.Styles;

namespace ExcelReportGenerator.ColumnarReport
{
    /// <summary>
    /// default styles for sheet
    /// </summary>
    public interface IExcelReportStyle:IExcelStyle
    {
        /// <summary>
        /// default height for all rows
        /// </summary>
        public double? DefaultRowHeight { get; set; }

        /// <summary>
        /// default width for all columns
        /// </summary>
        public double?  DefaultColumnWidth { get; set; }

        /// <summary>
        /// a collection of colors to use for a striped background color pattern for rows
        /// for example if you use [white,gray] as pallet, rows will be white and gray one after another
        /// </summary>
        public string[]? RowsColorPallet { get; set; }

        /// <summary>
        /// a collection of colors to use for a striped background color pattern for columns
        /// for example if you use [white,gray] as pallet, columns will be white and gray one after another
        /// </summary>
        public string[]? ColumnsColorPallet { get; set; }
    }
}
