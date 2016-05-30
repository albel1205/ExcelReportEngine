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
    public class Border : Attribute, IRangeDecorator
    {
        public ExcelBorderStyle Top { get; set; }
        public ExcelBorderStyle Right { get; set; }
        public ExcelBorderStyle Bottom { get; set; }
        public ExcelBorderStyle Left { get; set; }
        
        public void ApplyToSheet(ExcelWorksheet sheet, RangeInfo range)
        {
            var cell = sheet.Cells[range.FromRow, range.FromColum, range.ToRow, range.ToColumn];
            cell.Style.Border.Top.Style = Top;
            cell.Style.Border.Right.Style = Right;
            cell.Style.Border.Bottom.Style = Bottom;
            cell.Style.Border.Left.Style = Left;
        }
    }
}
