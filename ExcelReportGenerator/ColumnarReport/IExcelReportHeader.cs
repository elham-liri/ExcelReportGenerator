using ExcelReportGenerator.Styles;

namespace ExcelReportGenerator.ColumnarReport
{
    /// <summary>
    /// define header row properties
    /// </summary>
    public interface IExcelReportHeader
    {
        /// <summary>
        /// if true the header wil not be displayed
        /// </summary>
        public bool Disabled { get; set; }

        /// <summary>
        /// order of header in case there are more than one header row
        /// </summary>
        public int Order { get; set; }

        /// <summary>
        /// header row style; will overwrite columnStyle
        /// </summary>
        public IExcelStyle Style { get; set; }  

        /// <summary>
        /// if true header will be built upon columns
        /// </summary>
        public bool ColumnBased { get; set; }

        /// <summary>
        /// if ColumnBased=false then header will be built upon these parts
        /// </summary>
        public IEnumerable<IExcelHeaderPart> HeaderParts { get; set; }
    }
}
