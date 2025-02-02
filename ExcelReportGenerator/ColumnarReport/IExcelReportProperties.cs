namespace ExcelReportGenerator.ColumnarReport
{
    /// <summary>
    /// determine the general properties for an excel file
    /// </summary>
    public interface IExcelReportProperties
    {
        /// <summary>
        ///if true sets the sheet direction right to left
        /// </summary>
        public bool RightToLeft { get; set; }

        /// <summary>
        /// if true sets the sheet as protected
        /// </summary>
        public bool IsProtected { get; set; }
        
        /// <summary>
        ///set a password for sheet if it's protected 
        /// </summary>
        public string? ProtectionPassword  { get; set; }

        /// <summary>
        /// if true allow users to delete rows
        /// </summary>
        public bool AllowDeleteRows { get; set; }

        /// <summary>
        /// if true allow users to delete columns
        /// </summary>
        public bool AllowDeleteColumns { get; set; }

        /// <summary>
        /// if true allow users to select locked cells
        /// </summary>
        public bool AllowSelectLockedCells { get; set; }

        /// <summary>
        /// count of columns that should be frozen from left
        /// </summary>
        public int FrozenColumns { get; set; }

        /// <summary>
        /// count of rows that should be frozen from top
        /// </summary>
        public int FrozenRows { get; set; }

        /// <summary>
        /// if true ,all columns wil be set to be autofit
        /// </summary>
        public bool AutoFitColumns { get; set; }
    }
}
