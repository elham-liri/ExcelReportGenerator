using OfficeOpenXml.Style;

namespace ExcelReportGenerator.Styles
{
    /// <summary>
    /// determine border style
    /// </summary>
    public interface IExcelBorderStyle
    {
        /// <summary>
        /// if true, applies border properties to all sides
        /// </summary>
        public bool BorderAround { get; set; }

        /// <summary>
        /// border color for all sides- can be in one of these forms:
        /// example1 : black
        /// example2 : #000000
        /// example3 : rbg(0,0,0)
        /// example4 : rgba(0,0,0,1)
        /// </summary>
        public string? AroundColor { get; set; }

        /// <summary>
        /// border style for all sides
        /// </summary>
        public ExcelBorderStyle? AroundStyle { get; set; }

        /// <summary>
        /// border color for top border- can be in one of these forms:
        /// example1 : black
        /// example2 : #000000
        /// example3 : rbg(0,0,0)
        /// example4 : rgba(0,0,0,1)
        /// </summary>
        public string? TopColor { get; set; }
        
        /// <summary>
        /// border style for top border
        /// </summary>
        public ExcelBorderStyle? TopStyle { get; set; }

        /// <summary>
        /// border color for left border- can be in one of these forms:
        /// example1 : black
        /// example2 : #000000
        /// example3 : rbg(0,0,0)
        /// example4 : rgba(0,0,0,1)
        /// </summary>
        public string? LeftColor { get; set; }

        /// <summary>
        /// border style for left border
        /// </summary>
        public ExcelBorderStyle? LeftStyle { get; set; }

        /// <summary>
        /// border color for bottom border- can be in one of these forms:
        /// example1 : black
        /// example2 : #000000
        /// example3 : rbg(0,0,0)
        /// example4 : rgba(0,0,0,1)
        /// </summary>
        public string? BottomColor { get; set; }

        /// <summary>
        /// border style for bottom border
        /// </summary>
        public ExcelBorderStyle? BottomStyle { get; set; }

        /// <summary>
        /// border color for right border- can be in one of these forms:
        /// example1 : black
        /// example2 : #000000
        /// example3 : rbg(0,0,0)
        /// example4 : rgba(0,0,0,1)
        /// </summary>
        public string? RightColor { get; set; }

        /// <summary>
        /// border style for right border
        /// </summary>
        public ExcelBorderStyle? RightStyle { get; set; }

    }
}
