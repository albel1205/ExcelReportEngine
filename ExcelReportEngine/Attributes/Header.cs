using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReportEngine.Models;
using OfficeOpenXml;

namespace ExcelReportEngine.Attributes
{
    public class Header : AttributeBase
    {
        public int StartRow { get; set; }
        public int StartColumn { get; set; }

        public override void ApplyToSheet(ExcelWorksheet sheet, RangeInfo range, object value)
        {
            string[] headers = value as string[];

            int count = StartColumn;
            foreach(var header in headers)
            {
                sheet.Cells[StartRow, count, StartRow, count].Value = header;
                count++;
            }
        }
    }
}
