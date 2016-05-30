using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelReportEngine.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class Worksheet : Attribute, ISheetDecorator
    {
        public string Name { get; set; }
        public bool IsShowGridLine { get; set; }

        public void ApplyToSheet(ExcelWorksheet sheet)
        {
            sheet.Name = Name;
            sheet.View.ShowGridLines = IsShowGridLine;
        }
    }
}
