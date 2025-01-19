namespace ExcelReportGenerator.Interfaces
{
    public interface IExcelColumn
    {
        public string Name { get; set; } 
        public int Order { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public bool Required { get; set; }
        public bool IsHorizontal { get; set; }
        public bool Excluded { get; set; }
        public bool HasTotalSum { get; set; }

        public bool Locked { get; set; }
        public bool AddComment { get; set; }
        public bool IsDynamicColumn { get; set; }
        public IExcelColumnDataProperties DataProperties { get; set; }
    }
}
