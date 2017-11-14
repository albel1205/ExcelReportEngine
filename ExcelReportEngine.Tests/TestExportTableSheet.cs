using ExcelReportEngine.Tests.TestModels;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportEngine.Tests
{
    public class TestExportTableSheet : TestReportBase
    {
        private new string EXCEL_FILE = "report.xlsx";

        [TestCase]
        public void TestExportFileSuccess()
        {
            var report = GetSampleReport();
            var stream = report.GetReportStream();

            SaveExcelReport(stream);
            Assert.IsTrue(true);
        }

        private SampleExcelReport GetSampleReport()
        {
            var sheet = new SampleTableWorkSheet()
            {
                CustomerComplaintTable = GetSampleCustomerComplaintTable(),
            };
            var sheetArray = new SampleTableWorkSheet[] { sheet };

            var report = new SampleExcelReport(sheetArray);

            return report;
        }

        private SampleCustomerComplaintTable GetSampleCustomerComplaintTable()
        {
            string[] headers = new string[]
            {
                "Registered Date",
                "Complaint Type",
                "Complaint Name",
                "Address"
            };

            SampleCustomerComplaintItem[] items = new SampleCustomerComplaintItem[]
            {
                new SampleCustomerComplaintItem()
                {
                    RegisteredDate = new DateTime(2016, 5, 20),
                    ComplaintTypeText = "Complaint Text",
                    ComplaintName = "Testing Name",
                    Address = "Dong Nai"
                },
                new SampleCustomerComplaintItem()
                {
                    RegisteredDate = new DateTime(2016, 5, 21),
                    ComplaintTypeText = "Complaint Text",
                    ComplaintName = "Testing Name",
                    Address = "Ho Chi Minh"
                },
                new SampleCustomerComplaintItem()
                {
                    RegisteredDate = new DateTime(2016, 5, 22),
                    ComplaintTypeText = "Complaint Text",
                    ComplaintName = "Testing Name",
                    Address = "Da Nang"
                },
                new SampleCustomerComplaintItem()
                {
                    RegisteredDate = new DateTime(2016, 5, 23),
                    ComplaintTypeText = "Complaint Text",
                    ComplaintName = "Testing Name",
                    Address = "Ha Noi"
                },
                new SampleCustomerComplaintItem()
                {
                    RegisteredDate = new DateTime(2016, 5, 24),
                    ComplaintTypeText = "Complaint Text",
                    ComplaintName = "Testing Name",
                    Address = "Binh Duong"
                }
            };

            return new SampleCustomerComplaintTable(headers, items);
        }
    }
}
