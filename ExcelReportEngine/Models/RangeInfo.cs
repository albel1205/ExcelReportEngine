using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportEngine.Models
{
    public class RangeInfo
    {
        public RangeInfo()
        {

        }

        public RangeInfo(int fromRow, int fromCol, int toRow, int toCol)
        {
            this.FromRow = fromRow;
            this.FromColum = fromCol;
            this.ToRow = toRow;
            this.ToColumn = toCol;
        }

        public int FromRow { get; set; }
        public int FromColum { get; set; }
        public int ToRow { get; set; }
        public int ToColumn { get; set; }

        public string Value { get; set; }
    }
}
