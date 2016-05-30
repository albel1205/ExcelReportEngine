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
        public virtual ExcelWorksheet AddWorksheetTo(ExcelWorkbook workbook)
        {
            var worksheetAttr = Attribute.GetCustomAttribute(this.GetType(), typeof(Worksheet)) as Worksheet;
            var sheetName = worksheetAttr.Name;

            var worksheet = workbook.Worksheets.Add(sheetName);

            Range range = null;

            var props = this.GetType().GetProperties();
            foreach(var prop in props)
            {
                object[] attributes = prop.GetCustomAttributes(true);

                var cellsAttr = attributes.FirstOrDefault(x => x.GetType() == typeof(Cells) || x.GetType() == typeof(Cell));
                if (cellsAttr == null)
                {
                    continue;
                }
                else if(cellsAttr is Cells)
                {
                    var cellsAttribute = cellsAttr as Cells;
                    range = new Range(cellsAttribute.FromRow, cellsAttribute.FromColum,
                        cellsAttribute.ToRow, cellsAttribute.ToColumn);
                    range.Value = prop.GetValue(this).ToString();//check null
                }
                else if (cellsAttr is Cell)
                {
                    var cellsAttribute = cellsAttr as Cell;
                    range = new Range(cellsAttribute.Row, cellsAttribute.Column,
                        cellsAttribute.Row, cellsAttribute.Column);
                    range.Value = prop.GetValue(this).ToString();//check null
                }

                foreach (var attr in attributes)
                {
                    //if (!attr.GetType().IsSubclassOf(typeof(IRangeDecorator))) continue;

                    var attribute = attr as IRangeDecorator;
                    attribute.ApplyToSheet(worksheet, range);
                }
            }

            var sheetAttrs = this.GetType().GetCustomAttributes();
            foreach(var attr in sheetAttrs)
            {
                //if (!attr.GetType().IsSubclassOf(typeof(ISheetDecorator))) continue;

                var attribute = attr as ISheetDecorator;
                attribute.ApplyToSheet(worksheet);
            }

            return worksheet;
        }
    }
}
