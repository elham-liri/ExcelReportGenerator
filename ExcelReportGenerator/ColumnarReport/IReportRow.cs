namespace ExcelReportGenerator.ColumnarReport
{
    public interface IReportRow<T> where T : class, IReportCell ,new()
    {
        public IList<T> Cells { get; set; }
    }
}
