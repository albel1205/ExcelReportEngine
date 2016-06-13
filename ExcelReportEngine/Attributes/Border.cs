using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using ExcelReportEngine.Models;

namespace ExcelReportEngine.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class Border : AttributeBase
    {
        public ExcelBorderStyle Top { get; set; }
        public ExcelBorderStyle Right { get; set; }
        public ExcelBorderStyle Bottom { get; set; }
        public ExcelBorderStyle Left { get; set; }
        
        public override void ApplyToSheet(ExcelWorksheet sheet, RangeInfo range, object value)
        {
            var cell = sheet.Cells[range.FromRow, range.FromColumn, range.ToRow, range.ToColumn];
            cell.Style.Border.Top.Style = Top;
            cell.Style.Border.Right.Style = Right;
            cell.Style.Border.Bottom.Style = Bottom;
            cell.Style.Border.Left.Style = Left;

            base.ApplyToSheet(sheet, range, value);
        }
    }
}
