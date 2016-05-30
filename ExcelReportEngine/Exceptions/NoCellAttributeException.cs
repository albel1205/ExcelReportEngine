using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportEngine.Exceptions
{
    public class NoCellAttributeException : EntryPointNotFoundException
    {
        public NoCellAttributeException()
        {
        }

        public NoCellAttributeException(string message)
        : base(message)
        {
        }

        public NoCellAttributeException(string message, Exception inner)
        : base(message, inner)
        {
        }
    }
}
