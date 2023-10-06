using System.Drawing;
using OfficeOpenXml.Style;

namespace FileTools.Models.Xlsx
{
    public class CustomStyle
    {
        public CellFormat? CellFormat { get; set; }
        public int FontSize { get; set; }
        public bool FontBold { get; set; }
        public bool FontItalic { get; set; }
        public Color? FontColor { get; set; }
        public string? FontName { get; set; }
        public ExcelFillStyle? PatternType { get; set; }
        public Color? BackgroundColor { get; set; }
        public ExcelBorderStyle? BorderStyle { get; set; }
        public Color? BorderColor { get; set; }
        public ExcelHorizontalAlignment? HorizontalAlignment { get; set; }
        public ExcelVerticalAlignment? VerticalAlignment { get; set; }
    }
}