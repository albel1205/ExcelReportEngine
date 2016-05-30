using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using ExcelReportEngine.Models;

namespace ExcelReportEngine.Tests
{
    public interface IExcelReadable
    {
        ExcelRange GetRange(string path, Range range);
    }
}
