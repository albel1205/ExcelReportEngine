using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportEngine.Tests.TestModels
{
    public class SampleCustomerComplaintItem : TableItemBase
    {
        public DateTime RegisteredDate { get; set; }
        public string ComplaintTypeText { get; set; }
        public string ComplaintName { get; set; }
        public string Address { get; set; }
    }
}
