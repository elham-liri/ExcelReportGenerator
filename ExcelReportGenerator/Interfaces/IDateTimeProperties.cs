namespace ExcelReportGenerator.Interfaces
{
    public interface IDateTimeProperties
    {
        public bool ShowDate { get; set; }
        public bool ShowTime { get; set; }
        public string? Format { get; set; }
    }
}
