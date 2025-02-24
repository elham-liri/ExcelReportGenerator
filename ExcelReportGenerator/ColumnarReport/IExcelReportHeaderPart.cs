namespace ExcelReportGenerator.ColumnarReport
{
    /// <summary>
    /// define a header part
    /// </summary>
    public interface IExcelReportHeaderPart
    {
        /// <summary>
        /// order of header part
        /// </summary>
        public int Order { get; set; }

        /// <summary>
        /// colspan of header part
        /// </summary>
        public int Colspan { get; set; }

        /// <summary>
        /// the text that will be displayed
        /// </summary>
        public string Value { get; set; }
    }
}
