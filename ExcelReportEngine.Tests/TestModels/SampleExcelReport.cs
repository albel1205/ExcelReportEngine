using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReportEngine;
using ExcelReportEngine.Attributes;

namespace ExcelReportEngine.Tests.TestModels
{
    [Workbook]
    public class SampleExcelReport : ReportBase
    {
        public SampleExcelReport(SheetBase[] sheets)
            : base(sheets)
        {

        }
    }
}
