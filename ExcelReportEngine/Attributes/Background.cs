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
    public class Background : AttributeBase
    {
        public Color Color { get; set; }

        public override void ApplyToSheet(ExcelWorksheet sheet, RangeInfo range, object value)
        {
            var cell = sheet.Cells[range.FromRow, range.FromColumn, range.ToRow, range.ToColumn];
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(Color);

            base.ApplyToSheet(sheet, range, value);
        }
    }
}
