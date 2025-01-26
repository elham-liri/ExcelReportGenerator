namespace ExcelReportGenerator.Interfaces
{
    public interface ISingleSheetExcelRequest<T> where T : class,IExcelData
    {
        public string SheetName { get; set; }
        public List<T> DataList { get; set; }
        public IBaseExcelProfile ExcelProfile { get; set; }
    }
}
