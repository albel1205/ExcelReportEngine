using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReportEngine.Models;
using OfficeOpenXml;

namespace ExcelReportEngine.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class Array : AttributeBase
    {
        public override void ApplyToSheet(ExcelWorksheet sheet, RangeInfo range, object value)
        {
            base.ApplyToSheet(sheet, range, value);
        }
    }
}
