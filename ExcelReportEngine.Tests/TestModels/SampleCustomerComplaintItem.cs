using ExcelReportEngine.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportEngine.Tests.TestModels
{
    public class SampleCustomerComplaintItem : TableItemBase
    {
        [Cell(Column = 1)]
        public DateTime RegisteredDate { get; set; }
        [Cell(Column = 2)]
        public string ComplaintTypeText { get; set; }
        [Cell(Column = 3)]
        public string ComplaintName { get; set; }
        [Cell(Column = 4)]
        public string Address { get; set; }
    }
}
