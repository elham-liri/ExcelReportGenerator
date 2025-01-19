namespace ExcelReportGenerator.Interfaces
{
    public interface IBaseExcelProfile
    {
        List<IExcelColumn> Columns { get; set; }
        IExcelProfileDefaultProperties DefaultProperties { get; set; }
        IExcelProfileCellStyle CellStyle { get; set; }
        IExcelProfileHeaderStyle HeaderStyle { get; set; }
        IExcelProfileTotalSumStyle TotalSumStyle { get; set; }

    }
}
