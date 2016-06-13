using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReportEngine.Models;
using OfficeOpenXml;

namespace ExcelReportEngine.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class Font : AttributeBase
    {
        public string FontName { get; set; }
        public FontWeights FontWeight { get; set; }
        public float Size { get; set; }
        public int ColorRgb { get; set; }
        
        public override void ApplyToSheet(ExcelWorksheet sheet, RangeInfo range, object value)
        {
            var cell = sheet.Cells[range.FromRow, range.FromColumn, range.ToRow, range.ToColumn];
            cell.Style.Font.Name = FontName;
            cell.Style.Font.Size = Size;
            cell.Style.Font.Color.SetColor(Color.FromArgb(ColorRgb));

            if (FontWeight == FontWeights.Bold)
            {
                cell.Style.Font.Bold = true;
            }
            else if (FontWeight == FontWeights.Italic)
            {
                cell.Style.Font.Italic = true;
            }
            else if (FontWeight == FontWeights.Underscore)
            {
                cell.Style.Font.UnderLine = true;
            }

            base.ApplyToSheet(sheet, range, value);
        }
    }

    public enum FontWeights
    {
        Bold = 1,
        Italic = 2,
        Underscore = 3,
        Normal = 4
    }
}
