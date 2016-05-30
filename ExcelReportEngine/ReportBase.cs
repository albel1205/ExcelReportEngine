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
            throw new NotImplementedException();
        }
    }
}
