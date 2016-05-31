using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReportEngine.Tests.TestModels;
using ExcelReportEngine.Models;
using System.IO;
using ExcelReportEngine.Exceptions;
using OfficeOpenXml;

namespace ExcelReportEngine.Tests
{
    [TestFixture]
    public class TestCustomerComplaintSummaryReport : TestReportBase
    {
        private const string EXCEL_FILE = "CustomerComplaintReport.xlsx";

        [TestCase]
        public void TestExportFileSuccess()
        {
            var report = GetSampleExcelReport();
            var stream = report.GetReportStream();

            SaveExcelReport(stream);
            Assert.IsTrue(true);
        }

        private void SaveExcelReport(MemoryStream stream)
        {
            string filePath = @"c:\temp\" + EXCEL_FILE;
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            WriteToFile(stream, filePath);
        }

        private SampleExcelReport GetSampleExcelReport()
        {
            var sheet = new SampleWorksheet()
            {
                CustomerComplaintTable = GetSampleCustomerComplaintTable(),
                FromPeriod = "2015-15-01",
                FromPeriodLabel = "From:",
                ToPeriod = "2015-15-01",
                ToPeriodLabel = "To:",
                CustomerName = "Posten Aarben AB",
                CustomerNameLabel = "Customer:",
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

        [TestCase]
        public void TestNoCellAttributeException()
        {
            Assert.Throws<NoCellAttributeException>(TestNoCellAttributeMethod);
        }

        private void TestNoCellAttributeMethod()
        {
            var report = GetSampleNoCellAttributeReport();
            var workbook = report.GetReportStream();
        }

        private SampleExcelReport GetSampleNoCellAttributeReport()
        {
            var sheet = new SampleNoCellAttribute()
            {
                Title = "Customer Complaint Summary"
            };
            var sheetArray = new SampleNoCellAttribute[] { sheet };

            var report = new SampleExcelReport(sheetArray);

            return report;
        }

        [TestCase]
        public void TestPropertyNullException()
        {
            Assert.Throws<ArgumentNullException>(TestPropertyNullMethod);
        }

        private void TestPropertyNullMethod()
        {
            var report = GetSampleNullPropertyValueReport();
            var workbook = report.GetReportStream();
        }

        private SampleExcelReport GetSampleNullPropertyValueReport()
        {
            var sheet = new SampleWorksheet()
            {
                FromPeriod = "2015-15-01",
                FromPeriodLabel = "From:",
                ToPeriod = "2015-15-01",
                ToPeriodLabel = "To:",
                CustomerName = "Posten Aarben AB",
                CustomerNameLabel = "Customer:",
                Title = null
            };
            var sheetArray = new SampleWorksheet[] { sheet };

            var report = new SampleExcelReport(sheetArray);

            return report;
        }

        [TestCase]
        public void TestTitleProperty()
        {
            var report = GetSampleExcelReport();
            var stream = report.GetReportStream();

            using (ExcelPackage package = new ExcelPackage())
            {
                package.Load(stream);

                var workbook = package.Workbook;
                var sheet = workbook.Worksheets.First();
                Assert.IsNotNull(sheet);

                var sheetName = sheet.Name;
                var isShowGrid = sheet.View.ShowGridLines;

                var isMerged = sheet.Cells[2, 2, 2, 7].Merge;
                var value = sheet.Cells[2, 2, 2, 7].RichText.Text;
                var fontName = sheet.Cells[2, 2, 2, 7].Style.Font.Name;
                var fontSize = sheet.Cells[2, 2, 2, 7].Style.Font.Size;
                var isBold = sheet.Cells[2, 2, 2, 7].Style.Font.Bold;
                var alignment = sheet.Cells[2, 2, 2, 7].Style.HorizontalAlignment;

                Assert.IsTrue(sheetName.Equals("Sample Work Sheet"));
                Assert.IsTrue(isShowGrid == false);

                Assert.IsTrue(isMerged);
                Assert.IsTrue(value.Equals("Customer Complaint Summary"));
                Assert.IsTrue(fontName.Equals("Arial"));
                Assert.IsTrue(fontSize == 25);
                Assert.IsTrue(isBold);
                Assert.IsTrue(alignment == OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
            }
        }
    }
}
