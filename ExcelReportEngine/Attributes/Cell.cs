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
    public class Cell : AttributeBase, ILocatable
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public string Id { get; set; }
        public int MarginRow { get; set; }
        public int MarginColumn { get; set; }
        public string MarginWith { get; set; }

        public int EndRow {
            get
            {
                return Row + Width;
            }
        }

        public int EndColumn
        {
            get
            {
                return Column + Height;
            }
        }
        
        public override void ApplyToSheet(ExcelWorksheet sheet, RangeInfo range, object value)
        {
            sheet.Cells[Row == 0 ? range.FromRow : Row, Column == 0 ? range.FromColumn : Column
                , EndRow == 0 ? range.ToRow : EndRow, EndColumn == 0 ? range.ToColumn : EndColumn].Merge = true;
            sheet.Cells[Row == 0 ? range.FromRow : Row, Column == 0 ? range.FromColumn : Column
                , EndRow == 0 ? range.ToRow : EndRow, EndColumn == 0 ? range.ToColumn : EndColumn].Value = value;

            base.ApplyToSheet(sheet, range, value);
        }

        public virtual Range GetRange()
        {
            if (string.IsNullOrEmpty(MarginWith))
            {
                return new Range(Row, Column, EndRow, EndColumn);
            }

            Row = 0; Column = 0;
            return new Range(0, 0, EndRow, EndColumn);
        }
    }
}
