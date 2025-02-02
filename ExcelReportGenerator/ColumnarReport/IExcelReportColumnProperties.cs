namespace ExcelReportGenerator.ColumnarReport
{
    /// <summary>
    /// determine common properties to apply on column or all cells in a column
    /// </summary>
    public interface IExcelReportColumnProperties
    {
        /// <summary>
        /// if true , lock the cells in column for editing if the sheet is protected
        /// </summary>
        public bool Locked { get; set; }

        /// <summary>
        /// if true set the column width based on contents in the range
        /// </summary>
        public bool AutoFit { get; set; }

        /// <summary>
        /// if true the column will be hidden
        /// </summary>
        public bool Hidden { get; set; }

        /// <summary>
        /// if true wrap the text in cells in the range 
        /// </summary>
        public bool WrapText { get; set; }

        /// <summary>
        /// format to determine how to display value of the cells in the range 
        /// </summary>
        public string? FormatCell { get; set; }
    }
}
