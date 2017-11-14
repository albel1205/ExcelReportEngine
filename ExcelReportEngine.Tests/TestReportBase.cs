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
        public string EXCEL_FILE = "report.xlsx";

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

        protected void SaveExcelReport(MemoryStream stream)
        {
            string filePath = @"c:\temp\" + EXCEL_FILE;
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            WriteToFile(stream, filePath);
        }
    }
}
