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
    public class Margin : AttributeBase
    {
        public string MarginWith { get; set; }
        public MarginKinds MarginKind { get; set; }
        public int NumberOfCells { get; set; }

        public override void ApplyToSheet(ExcelWorksheet sheet, RangeInfo range)
        {
            if(this.MarginKind == MarginKinds.Row)
            {
                range.FromRow += NumberOfCells;
            }
            else if(this.MarginKind == MarginKinds.Column)
            {
                range.FromColumn += NumberOfCells;
            }

            base.ApplyToSheet(sheet, range);
        }
    }

    public enum MarginKinds
    {
        Row = 1,
        Column = 2
    }
}
