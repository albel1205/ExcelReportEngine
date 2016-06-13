using ExcelReportEngine.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReportEngine.Models;
using OfficeOpenXml;

namespace ExcelReportEngine
{
    public abstract class TableBase : AttributeBase
    {
        public TableBase(string[] headers, TableItemBase[] data)
        {
            this.Headers = headers;
            this.Data = data;
        }

        public string[] Headers { get; set; }
        public TableItemBase[] Data { get; set; }

        public override void ApplyToSheet(ExcelWorksheet sheet, RangeInfo range, object value)
        {
            base.ApplyToSheet(sheet, range,value);
        }
    }
}
