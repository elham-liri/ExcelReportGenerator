using ExcelReportGenerator.Interfaces;

namespace ExcelReportGenerator.ExcelEntities
{
    public class NumberProperties:INumberProperties
    {
        public bool UseGroupingSeparator { get; set; } = false;
        public string? Format { get; set; }
    }
}
