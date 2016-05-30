using ExcelReportEngine.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportEngine.Attributes
{
    public interface IRangeDecorator
    {
        void ApplyToSheet(ExcelWorksheet sheet, Range range);
    }
}
