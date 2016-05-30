using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelReportEngine.Attributes
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public class Row : Attribute, ISheetDecorator
    {
        public double Height { get; set; }
        public int Index { get; set; }

        public void ApplyToSheet(ExcelWorksheet sheet)
        {
            sheet.Column(Index).Width = Height;
        }
    }
}
