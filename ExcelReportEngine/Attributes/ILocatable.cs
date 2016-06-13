using ExcelReportEngine.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportEngine.Attributes
{
    public interface ILocatable
    {
        Range GetRange();
    }
}
