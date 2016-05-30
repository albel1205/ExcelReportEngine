using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportEngine
{
    public abstract class ReportBase : IExportable
    {
        public SheetBase[] Sheets { get; set; }

        public ReportBase(SheetBase[] sheets)
        {
            this.Sheets = sheets;
        }

        public MemoryStream GetReportStream()
        {
            var ms = new MemoryStream();

            using (var package = new ExcelPackage(ms))
            {
                var workbook = package.Workbook;
                foreach(var sheet in Sheets)
                {
                    sheet.AddWorksheetTo(workbook);
                }

                package.Save();
                package.Dispose();
            }

            return ms;
        }
    }
}
