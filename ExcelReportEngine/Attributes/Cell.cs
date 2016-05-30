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
    public class Cell : Attribute, IRangeDecorator, ILocatable
    {
        public int Row { get; set; }
        public int Column { get; set; }
        
        public virtual void ApplyToSheet(ExcelWorksheet sheet, RangeInfo range)
        {
            sheet.Cells[Row, Column].Merge = true;
            sheet.Cells[Row, Column].Value = range.Value;
        }

        public virtual RangeInfo GetRange()
        {
            return new RangeInfo(Row, Column, Row, Column);
        }
    }
}
