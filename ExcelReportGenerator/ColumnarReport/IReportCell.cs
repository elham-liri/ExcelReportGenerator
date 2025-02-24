namespace ExcelReportGenerator.ColumnarReport
{
    public interface IReportCell
    {
        public int Order { get; set; }
        public object? Value { get; set; }
    }
}
