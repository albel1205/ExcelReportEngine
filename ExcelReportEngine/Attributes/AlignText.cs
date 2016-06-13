using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReportEngine.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelReportEngine.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class AlignText : AttributeBase
    {
        public ExcelHorizontalAlignment Align { get; set; }

        public override void ApplyToSheet(ExcelWorksheet sheet, RangeInfo range, object value)
        {
            var cell = sheet.Cells[range.FromRow, range.FromColumn, range.ToRow, range.ToColumn];
            cell.Style.HorizontalAlignment = Align;

            base.ApplyToSheet(sheet, range,value);
        }
    }
}
