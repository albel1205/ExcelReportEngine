using ExcelReportEngine.Attributes;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportEngine.Tests.TestModels
{
    public class SampleCustomerComplaintTable : TableBase
    {
        [Header(StartRow = 1, StartColumn = 1)]
        public new string[] Headers { get; set; }

        [TableData(StartRow = 2)]
        public new TableItemBase[] Data { get; set; }

        public SampleCustomerComplaintTable(string[] headers, SampleCustomerComplaintItem[] data)
            : base(headers, data)
        {
            this.Headers = headers;
            this.Data = data;
        }
    }
}
