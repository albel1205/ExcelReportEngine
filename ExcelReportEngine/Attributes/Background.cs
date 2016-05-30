using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReportEngine.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelReportEngine.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class Background : Attribute, IRangeDecorator
    {
        public Color Color { get; set; }

        public void ApplyToSheet(ExcelWorksheet sheet, Range range)
        {
            var cell = sheet.Cells[range.FromRow, range.FromColum, range.ToRow, range.ToColumn];
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(Color);
        }
    }
}
