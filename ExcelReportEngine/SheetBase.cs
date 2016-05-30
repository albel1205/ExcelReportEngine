using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using ExcelReportEngine.Attributes;
using ExcelReportEngine.Models;
using System.Reflection;

namespace ExcelReportEngine
{
    public abstract class SheetBase
    {
        public virtual ExcelWorksheet GetWorkSheet(ExcelWorkbook workbook)
        {
            var worksheetAttr = Attribute.GetCustomAttribute(this.GetType(), typeof(Worksheet)) as Worksheet;
            var sheetName = worksheetAttr.Name;

            var worksheet = workbook.Worksheets.Add(sheetName);

            Range range = null;

            var props = this.GetType().GetProperties();
            foreach(var prop in props)
            {
                object[] attributes = prop.GetCustomAttributes(true);

                var cellsAttr = attributes.FirstOrDefault(x => x.GetType() == typeof(Cells));
                if (cellsAttr == null)
                {
                    continue;
                }
                else
                {
                    var cellsAttribute = cellsAttr as Cells;
                    range = new Range(cellsAttribute.FromRow, cellsAttribute.FromColum,
                        cellsAttribute.ToRow, cellsAttribute.ToColumn);
                }

                foreach(var attr in attributes)
                {
                    if (attr.GetType() != typeof(IRangeDecorator)) continue;

                    var attribute = attr as IRangeDecorator;
                    attribute.ApplyToSheet(worksheet, range);
                }
            }

            var sheetAttrs = this.GetType().GetCustomAttributes();
            foreach(var attr in sheetAttrs)
            {
                if (attr.GetType() != typeof(ISheetDecorator)) continue;

                var attribute = attr as ISheetDecorator;
                attribute.ApplyToSheet(worksheet);
            }

            return worksheet;
        }
    }
}
