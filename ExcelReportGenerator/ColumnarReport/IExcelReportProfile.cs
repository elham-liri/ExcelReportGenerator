namespace ExcelReportGenerator.ColumnarReport
{
    /// <summary>
    /// define a profile that an excel sheet will be produced based on it
    /// </summary>
    public interface IExcelReportProfile
    {
        /// <summary>
        /// list of columns
        /// </summary>
        public IEnumerable<IExcelReportColumn> Columns { get; set; }

        /// <summary>
        /// list of header rows
        /// </summary>
        public IEnumerable<IExcelReportHeader> Headers { get; set; }

        /// <summary>
        /// default properties for excel sheet
        /// </summary>
        public IExcelReportProperties DefaultProperties { get; set; }

        /// <summary>
        /// default style for all cells; can be overwritten by columnStyle , rowStyle and cellStyle
        /// </summary>
        public IExcelReportStyle DefaultStyle { get; set; }
    }
}
