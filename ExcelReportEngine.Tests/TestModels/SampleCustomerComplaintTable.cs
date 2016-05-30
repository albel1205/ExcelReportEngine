using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportEngine.Tests.TestModels
{
    public class SampleCustomerComplaintTable : TableBase<SampleCustomerComplaintItem>
    {
        public SampleCustomerComplaintTable(string[] headers, SampleCustomerComplaintItem[] data)
            : base(headers, data)
        {

        }
    }
}
