using OfficeOpenXml.Style;

namespace ExcelReportGenerator.Styles
{
    /// <summary>
    /// determine font style
    /// </summary>
    public interface IExcelFontStyle
    {
        /// <summary>
        /// font name
        /// </summary>
        public string? FontName { get; set; }

        /// <summary>
        /// font size
        /// </summary>
        public float? FontSize { get; set; }

        /// <summary>
        /// should font be bold
        /// </summary>
        public bool FontBold { get; set; }

        /// <summary>
        /// should fond be italic
        /// </summary>
        public bool FontItalic { get; set; }

        /// <summary>
        /// should font be underlined
        /// </summary>
        public bool FontUnderLine { get; set; }

        /// <summary>
        /// font color - can be in one of these forms:
        /// example1 : black
        /// example2 : #000000
        /// example3 : rbg(0,0,0)
        /// example4 : rgba(0,0,0,1)
        /// </summary>
        public string? FontColor { get; set; }

        /// <summary>
        /// horizontal alignment
        /// </summary>
        public ExcelHorizontalAlignment? HorizontalAlignment { get; set; }

        /// <summary>
        /// vertical alignment
        /// </summary>
        public ExcelVerticalAlignment? VerticalAlignment { get; set; }
    }
}
