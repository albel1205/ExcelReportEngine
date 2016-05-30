using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelReportEngine.Tests
{
    public interface IExcelWritable
    {
        void WriteToFile(MemoryStream excelStream, string path);
    }
}
