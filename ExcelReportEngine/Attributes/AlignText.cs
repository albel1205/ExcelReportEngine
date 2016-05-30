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
    public class AlignText : Attribute, IRangeDecorator
    {
        public ExcelHorizontalAlignment Align { get; set; }

        public void ApplyToSheet(ExcelWorksheet sheet, Range range)
        {
            var cell = sheet.Cells[range.FromRow, range.FromColum, range.ToRow, range.ToColumn];
            cell.Style.HorizontalAlignment = Align;
        }
    }
}
