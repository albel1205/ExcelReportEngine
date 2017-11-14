using ExcelReportEngine.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportEngine.Tests.TestModels
{
    [Worksheet(Name = "Sample Work Sheet", IsShowGridLine = true)]
    public class SampleTableWorkSheet : SheetBase
    {
        [Table]
        public SampleCustomerComplaintTable CustomerComplaintTable { get; set; }
    }
}
