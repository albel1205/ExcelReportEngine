using ExcelReportEngine.Attributes;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportEngine.Tests.TestModels
{
    [Worksheet(Name = "Sample Work Sheet", IsShowGridLine = false)]
    [Column(Index = 2, Width = 100)]
    [Column(Index = 4, Width = 20)]
    [Row(Index = 3, Height = 3)]
    public class SampleNoCellAttribute : SheetBase
    {
        [Font(ColorRgb = 0, FontName = "Arial", FontWeight = FontWeights.Bold, Size = 25)]
        [AlignText(Align = ExcelHorizontalAlignment.Center)]
        public string Title { get; set; }
    }
}
