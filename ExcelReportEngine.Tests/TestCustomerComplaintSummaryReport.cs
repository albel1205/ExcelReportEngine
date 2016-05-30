﻿using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReportEngine.Tests.TestModels;
using ExcelReportEngine.Models;

namespace ExcelReportEngine.Tests
{
    public class TestCustomerComplaintSummaryReport : TestReportBase
    {
        private const string EXCEL_FILE = "CustomerComplaintReport.xlsx";

        [Test]
        public void TestExport()
        {
            var report = GetSampleExcelReport();
            var stream = report.GetReportStream();

            WriteToFile(stream, EXCEL_FILE);

            var titleRange = GetRange(EXCEL_FILE, new Range(1,1,1,5));
            Assert.IsTrue(titleRange.Text.Trim().Equals("Complaint Summary Customer"));
            Assert.IsTrue(titleRange.Merge);
        }

        private SampleExcelReport GetSampleExcelReport()
        {
            var sheet = new SampleWorksheet()
            {
                //CustomerComplaintTable = GetSampleCustomerComplaintTable(),
                FromPeriod = "2015-15-01",
                ToPeriod = "2015-15-01",
                CustomerName = "Posten Aarben AB",
                Title = "Customer Complaint Summary"
            };
            var sheetArray = new SampleWorksheet[] { sheet };

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
