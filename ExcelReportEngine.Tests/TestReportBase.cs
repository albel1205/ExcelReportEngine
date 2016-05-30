using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using ExcelReportEngine.Models;

namespace ExcelReportEngine.Tests
{
    public abstract class TestReportBase : IExcelWritable, IExcelReadable
    {
        public ExcelRange GetRange(string path, Range range)
        {
            throw new NotImplementedException();
        }

        public void WriteToFile(MemoryStream excelStream, string path)
        {
            throw new NotImplementedException();
        }
    }
}
