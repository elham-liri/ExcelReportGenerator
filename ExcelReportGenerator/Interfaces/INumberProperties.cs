namespace ExcelReportGenerator.Interfaces
{
    public interface INumberProperties
    {
        public bool UseGroupingSeparator { get; set; }
        public string? Format { get; set; }
    }
}
