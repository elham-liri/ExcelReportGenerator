using ExcelReportGenerator.Interfaces;

namespace ExcelReportGenerator.ExcelEntities
{
    internal class DateTimeProperties:IDateTimeProperties
    {
        public bool ShowDate { get; set; } = true;

        public bool ShowTime { get; set; } = true;
        public string? Format { get; set; } = "mm/dd/yyyy h:mm";
    }
}
