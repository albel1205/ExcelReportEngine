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

        public RangeInfo(int fromRow, int fromCol, int toRow, int toCol, string value)
        {
            this.FromRow = fromRow;
            this.FromColumn = fromCol;
            this.ToRow = toRow;
            this.ToColumn = toCol;
        }

        public RangeInfo(Range range)
        {
            this.FromRow = range.FromRow;
            this.FromColumn = range.FromColum;
            this.ToRow = range.ToRow;
            this.ToColumn = range.ToColumn;
        }

        public int FromRow { get; set; }
        public int FromColumn { get; set; }
        public int ToRow { get; set; }
        public int ToColumn { get; set; }
    }
}
