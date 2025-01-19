namespace ExcelReportGenerator.Interfaces
{
    public interface IExcelColumnDataProperties
    {
        public IDateTimeProperties DateTimeProperties { get; set; }
        public string? CellFormat { get; set; }
        public INumberProperties NumberProperties { get; set; }
    }
}
