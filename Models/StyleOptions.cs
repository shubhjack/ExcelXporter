using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelXporter.Models
{
    public class StyleOptions
    {
        public HeaderStyle HeaderStyle { get; set; } = new();
        public ExcelCellStyle DefaultCellStyle { get; set; } = new();
        public BorderStyle BorderStyle { get; set; } = new();
    }

    public class HeaderStyle
    {
        public string BackgroundColorHex { get; set; } = string.Empty;
        public string FontColorHex { get; set; } = "000000"; 
    }

    public class ExcelCellStyle
    {
        public string FontColorHex { get; set; } = "000000"; // black
        public TextAlignment HorizontalAlignment { get; set; } = TextAlignment.Left; // or Center, Right
    }

    public enum TextAlignment
    {
        Center,
        Right,
        Left,
    }

    public class BorderStyle
    {
        public bool ApplyBorders { get; set; } = false;
        public string BorderColorHex { get; set; } = "000000";
        public BorderStyleValues Style { get; set; } = BorderStyleValues.Thin;
    }

}
