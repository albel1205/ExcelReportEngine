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
    [Worksheet(Name = "Sample Work Sheet", IsShowGridLine = false)]
    [Column(Index = 2, Width = 100)]
    [Column(Index = 4, Width = 20)]
    [Row(Index = 3, Height = 3)]
    public class SampleWorksheet : SheetBase
    {
        [Cells(Row = 2, Column = 2, ToRow = 2, ToColumn = 7)]
        [Font(ColorRgb = 0, FontName = "Arial", FontWeight = FontWeights.Bold, Size = 25)]
        [AlignText(Align = ExcelHorizontalAlignment.Center)]
        public string Title { get; set; }

        [Cell(Row = 4, Column = 2)]
        [Font(ColorRgb = 0, FontName = "Arial", FontWeight = FontWeights.Bold, Size = 13)]
        [AlignText(Align = ExcelHorizontalAlignment.Right)]
        public string FromPeriodLabel { get; set; }

        [Cell(Row = 4, Column = 3)]
        [Font(ColorRgb = 0, FontName = "Arial", FontWeight = FontWeights.Normal, Size = 13)]
        [AlignText(Align = ExcelHorizontalAlignment.Left)]
        public string FromPeriod { get; set; }

        [Cell(Row = 4, Column = 5)]
        [Font(ColorRgb = 0, FontName = "Arial", FontWeight = FontWeights.Bold, Size = 13)]
        [AlignText(Align = ExcelHorizontalAlignment.Right)]
        public string ToPeriodLabel { get; set; }

        [Cell(Row = 4, Column = 6)]
        [Font(ColorRgb = 0, FontName = "Arial", FontWeight = FontWeights.Normal, Size = 13)]
        [AlignText(Align = ExcelHorizontalAlignment.Left)]
        public string ToPeriod { get; set; }

        [Cell(Row = 5, Column = 2)]
        [Font(ColorRgb = 0, FontName = "Arial", FontWeight = FontWeights.Bold, Size = 13)]
        [AlignText(Align = ExcelHorizontalAlignment.Right)]
        public string CustomerNameLabel { get; set; }

        [Cell(Row = 5, Column = 3)]
        [Font(ColorRgb = 0, FontName = "Arial", FontWeight = FontWeights.Normal, Size = 13)]
        [AlignText(Align = ExcelHorizontalAlignment.Left)]
        public string CustomerName { get; set; }

        //[Table]
        //public SampleCustomerComplaintTable CustomerComplaintTable { get; set; }
    }
}
