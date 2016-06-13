using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReportEngine.Models;
using OfficeOpenXml;

namespace ExcelReportEngine.Attributes
{
    public abstract class AttributeBase : Attribute, IRangeDecorator
    {
        protected int _lastRow;
        protected int _lastColumn;

        public virtual void ApplyToSheet(ExcelWorksheet sheet, RangeInfo range, object value)
        {
            _lastRow = range.ToRow >= range.FromRow ? range.ToRow : range.FromRow;
            _lastColumn = range.ToColumn >= range.FromColumn ? range.ToColumn : range.FromColumn;
        }

        public virtual int GetLastColumn() { return 0; }

        public virtual int GetLastRow() { return 0; }
    }
}
