using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportEngine
{
    public abstract class TableBase<T>
    {
        public TableBase(string[] headers, T[] data)
        {
            this.Headers = headers;
            this.Data = data;
        }

        public string[] Headers { get; set; }
        public T[] Data { get; set; }
    }
}
