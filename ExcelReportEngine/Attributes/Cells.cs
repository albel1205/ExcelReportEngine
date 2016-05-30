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
    public class Cells : Attribute, IRangeDecorator
    {
        public int FromRow { get; set; }
        public int FromColum { get; set; }
        public int ToRow { get; set; }
        public int ToColumn { get; set; }
        
        public void ApplyToSheet(ExcelWorksheet sheet, Range range)
        {
            throw new NotImplementedException();
        }
    }
}
