using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReportEngine.Attributes;
using ExcelReportEngine.Models;
using OfficeOpenXml.Style;

namespace ExcelReportEngine.Tests.TestModels
{
    [Worksheet(Name = "Sample Work Sheet")]
    public class SampleWorksheet : SheetBase
    {
        [Cells(FromRow = 2, FromColum = 1, ToRow = 2, ToColumn = 5)]
        [Font(ColorRgb = 0, FontName = "Arial", FontWeight = FontWeights.Bold, Size = 25)]
        [AlignText(Align = ExcelHorizontalAlignment.Center)]
        public string Title { get; set; }

        [Cell(Row = 4, Column = 1)]
        [Font(ColorRgb = 0, FontName = "Arial", FontWeight = FontWeights.Bold, Size = 13)]
        [AlignText(Align = ExcelHorizontalAlignment.Right)]
        public string FromPeriodLabel { get; set; }

        [Cell(Row = 4, Column = 2)]
        [Font(ColorRgb = 0, FontName = "Arial", FontWeight = FontWeights.Normal, Size = 13)]
        [AlignText(Align = ExcelHorizontalAlignment.Left)]
        public string FromPeriod { get; set; }

        [Cell(Row = 4, Column = 4)]
        [Font(ColorRgb = 0, FontName = "Arial", FontWeight = FontWeights.Bold, Size = 13)]
        [AlignText(Align = ExcelHorizontalAlignment.Right)]
        public string ToPeriodLabel { get; set; }

        [Cell(Row = 4, Column = 5)]
        [Font(ColorRgb = 0, FontName = "Arial", FontWeight = FontWeights.Normal, Size = 13)]
        [AlignText(Align = ExcelHorizontalAlignment.Left)]
        public string ToPeriod { get; set; }

        [Cell(Row = 5, Column = 1)]
        [Font(ColorRgb = 0, FontName = "Arial", FontWeight = FontWeights.Bold, Size = 13)]
        [AlignText(Align = ExcelHorizontalAlignment.Right)]
        public string CustomerNameLabel { get; set; }

        [Cell(Row = 5, Column = 2)]
        [Font(ColorRgb = 0, FontName = "Arial", FontWeight = FontWeights.Normal, Size = 13)]
        [AlignText(Align = ExcelHorizontalAlignment.Left)]
        public string CustomerName { get; set; }

        //[Table]
        //public SampleCustomerComplaintTable CustomerComplaintTable { get; set; }
    }
}
