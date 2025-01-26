namespace ExcelReportGenerator.Interfaces
{
    public interface IBaseExcelProfile
    {
        List<IExcelReportColumn> Columns { get; set; }
        IExcelProfileDefaultProperties? DefaultProperties { get; set; }
        IExcelProfileCellStyle? CellStyle { get; set; }
        IExcelProfileHeaderStyle? HeaderStyle { get; set; }
        IExcelProfileTotalSumStyle? TotalSumStyle { get; set; }

        void InitializeColumns();
        void ResetColumns(IEnumerable<IExcelReportColumn> columns);
        void SetDefaultProperties();
        void SetHeaderStyle();
        void SetTotalSumStyle();
        void SetCellStyle();
        void AddDynamicColumns(IEnumerable<IExcelReportColumn> dynamicColumns);
        List<IExcelReportColumn> GetDisplayedVerticalColumns();
        List<IExcelReportColumn> GetDisplayedHorizontalColumns();
    }
}
