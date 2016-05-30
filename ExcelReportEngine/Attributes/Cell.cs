using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReportEngine.Models;
using OfficeOpenXml;

namespace ExcelReportEngine.Attributes
{
    [AttributeUsage( AttributeTargets.Property )]
    public class Cell : Attribute, IRangeDecorator
    {
        public int Row { get; set; }
        public int Column { get; set; }
        
        public void ApplyToSheet(ExcelWorksheet sheet, Range range)
        {
            sheet.Cells[Row, Column].Merge = true;
            sheet.Cells[Row, Column].Value = range.Value;
        }
    }
}
