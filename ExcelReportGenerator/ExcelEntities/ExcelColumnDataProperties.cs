using ExcelReportGenerator.Interfaces;

namespace ExcelReportGenerator.ExcelEntities
{
    public class ExcelColumnDataProperties:IExcelColumnDataProperties
    {
        public ExcelColumnDataProperties()
        {
            DateTimeProperties=new DateTimeProperties();
            NumberProperties=new NumberProperties();
        }

        public IDateTimeProperties DateTimeProperties { get; set; }
        public INumberProperties NumberProperties { get; set; }
        public string? CellFormat { get; set; }
    }
}
