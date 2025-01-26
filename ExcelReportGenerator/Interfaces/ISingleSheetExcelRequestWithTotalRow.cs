namespace ExcelReportGenerator.Interfaces
{
    public interface ISingleSheetExcelRequestWithTotalRow<T, TN> : ISingleSheetExcelRequest<T>
        where T : class, IExcelData 
        where TN : class, IExcelTotalDataModel
    {
        public TN TotalDataModel { get; set; }
    }
}
