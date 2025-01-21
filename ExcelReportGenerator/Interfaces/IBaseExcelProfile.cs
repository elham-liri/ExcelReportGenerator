namespace ExcelReportGenerator.Interfaces
{
    public interface IBaseExcelProfile
    {
        List<IExcelColumn> Columns { get; set; }
        IExcelProfileDefaultProperties? DefaultProperties { get; set; }
        IExcelProfileCellStyle? CellStyle { get; set; }
        IExcelProfileHeaderStyle? HeaderStyle { get; set; }
        IExcelProfileTotalSumStyle? TotalSumStyle { get; set; }

        void InitializeColumns();
        void ResetColumns(IEnumerable<IExcelColumn> columns);
        void SetDefaultProperties();
        void SetHeaderStyle();
        void SetTotalSumStyle();
        void SetCellStyle();
        void AddDynamicColumns(IEnumerable<IExcelColumn> dynamicColumns);
        List<IExcelColumn> GetDisplayedVerticalColumns();
        List<IExcelColumn> GetDisplayedHorizontalColumns();
    }
}
