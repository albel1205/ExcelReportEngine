using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelReportEngine.Attributes
{
    public abstract class AttributeBase : Attribute
    {
        public abstract void ApplyToSheet(ExcelWorksheet sheet);
    }
}
