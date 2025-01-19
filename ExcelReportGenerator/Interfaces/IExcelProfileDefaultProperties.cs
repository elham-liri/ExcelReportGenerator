namespace ExcelReportGenerator.Interfaces
{
    public interface IExcelProfileDefaultProperties
    {
        public double DefaultRowHeight { get; set; }
        public bool RightToLeft { get; set; }
        public bool AutoFitColumns { get; set; }
        public int FrozenColumns { get; set; }
        public int FrozenRows { get; set; }

        public float DefaultFontSize { get; set; }
        public string DefaultFontName { get; set; }

        public bool IsProtected { get; set; }
        public bool AllowDeleteRows { get; set; }
        public bool AllowDeleteColumns { get; set; }
        public bool AllowSelectLockedCells { get; set; }
    }
}
