using ExcelReportEngine.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelReportEngine.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class Cells : Cell
    {
        public int ToRow { get; set; }
        public int ToColumn { get; set; }
        
        public override void ApplyToSheet(ExcelWorksheet sheet, RangeInfo range)
        {
            sheet.Cells[Row, Column, ToRow, ToColumn].Merge = true;
            sheet.Cells[Row, Column, ToRow, ToColumn].Value = range.Value;
        }

        public override RangeInfo GetRange()
        {
            return new RangeInfo(Row, Column, ToRow, ToColumn);
        }
    }
}
