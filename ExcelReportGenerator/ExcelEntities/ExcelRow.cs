namespace ExcelReportGenerator.ExcelEntities
{
    public class ExcelRow
    {
        public ExcelRow()
        {
            Cells = new List<ExcelCell>();
            HorizontalCells = new List<ExcelCell>();
        }

        public object? Id { get; set; }
        public List<ExcelCell> Cells { get; set; }
        public List<ExcelCell> HorizontalCells { get; set; }
    }
}
