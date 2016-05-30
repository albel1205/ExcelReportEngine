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
        public ExcelRange GetRange(string path, RangeInfo range)
        {
            throw new NotImplementedException();
        }

        public void WriteToFile(MemoryStream excelStream, string path)
        {
            using (var filestream = new FileStream(path, FileMode.Create, FileAccess.ReadWrite))
            {
                excelStream.WriteTo(filestream);
                filestream.Close();
                excelStream.Close();
            }
        }
    }
}
