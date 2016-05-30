using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportEngine.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class Worksheet : Attribute
    {
        public string Name { get; set; }
        public bool IsShowGridLine { get; set; }
    }
}
