using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelReportEngine.Attributes
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public class Column : Attribute, ISheetDecorator
    {
        public double Width { get; set; }
        public int Index { get; set; }

        public void ApplyToSheet(ExcelWorksheet sheet)
        {
            sheet.Column(Index).Width = Width;
        }
    }
}
