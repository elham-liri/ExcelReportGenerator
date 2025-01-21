namespace ExcelReportGenerator.ExcelEntities
{
    public class ExcelReportRow
    {
        public ExcelReportRow()
        {
            Cells = new List<ExcelReportCell>();
            HorizontalCells = new List<ExcelReportCell>();
        }

        public object? Id { get; set; }
        public List<ExcelReportCell> Cells { get; set; }
        public List<ExcelReportCell> HorizontalCells { get; set; }
    }
}
