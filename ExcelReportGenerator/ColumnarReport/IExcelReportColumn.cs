namespace ExcelReportGenerator.ColumnarReport
{
    /// <summary>
    /// Use it to define a report column
    /// </summary>
    public interface IExcelReportColumn
    {
        /// <summary>
        /// property from which data of this column will be extracted
        /// </summary>
        public string SourceName { get; set; }

        /// <summary>
        /// column header
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// column order
        /// </summary>
        public int Order { get; set; }
        
        /// <summary>
        /// should column be horizontal. the horizontal column will be displayed in one row below other columns
        /// </summary>
        public bool IsHorizontal { get; set; }

        /// <summary>
        /// if true means column contains formula
        /// </summary>
        public bool NeedsCalculation { get; set; }

        /// <summary>
        /// if true means there must be a sum formula at the end of the column
        /// </summary>
        public bool ShowTotalSum { get; set; }

        /// <summary>
        /// column properties
        /// </summary>
        public IExcelReportColumnProperties ColumnProperties { get; set; }

        /// <summary>
        /// column style - this style will overwrite profile style
        /// </summary>
        public IExcelReportColumnStyle ColumnStyle { get; set; }

    }
}
