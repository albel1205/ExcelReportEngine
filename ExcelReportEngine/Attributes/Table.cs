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
    public class Table : Attribute, IRangeDecorator
    {
        public void ApplyToSheet(ExcelWorksheet sheet, RangeInfo range)
        {
            throw new NotImplementedException();
        }
    }
}
