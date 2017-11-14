using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReportEngine.Models;
using OfficeOpenXml;
using System.Reflection;
using ExcelReportEngine.Exceptions;

namespace ExcelReportEngine.Attributes
{
    public class TableData : AttributeBase
    {
        public int StartRow { get; set; }

        private Type iLocatableType = typeof(ILocatable);
        private Type iRangeDecoratorType = typeof(IRangeDecorator);

        public override void ApplyToSheet(ExcelWorksheet sheet, RangeInfo range, object value)
        {
            TableItemBase[] items = value as TableItemBase[];

            int rowCount = StartRow;
            foreach(var item in items)
            {
                var props = item.GetType().GetProperties();
                foreach (var prop in props)
                {
                    var propAttributes = GetPropertyAttributes(prop);

                    RangeInfo rangeInfo = GetRangeObject(propAttributes.ToArray(), prop);
                    object propValue = prop.GetValue(item);

                    rangeInfo.FromRow = rowCount;
                    rangeInfo.ToRow = rowCount;
                    propAttributes.ForEach(x => x.ApplyToSheet(sheet, rangeInfo, propValue));
                }

                rowCount++;
            }
        }

        private List<IRangeDecorator> GetPropertyAttributes(PropertyInfo property)
        {
            return property.GetCustomAttributes(true)
                    .Where(x => iRangeDecoratorType.IsInstanceOfType(x))
                    .Cast<IRangeDecorator>()
                    .ToList();
        }

        private RangeInfo GetRangeObject(object[] attributes, PropertyInfo prop)
        {
            var cellAttr = attributes.FirstOrDefault(x => iLocatableType.IsInstanceOfType(x));
            if (cellAttr == null)
            {
                throw new NoCellAttributeException(
                    string.Format("There is no Cell / Cells attribute on the {0} property", prop.Name));
            }

            var range = (cellAttr as ILocatable).GetRange();

            var rangeInfo = new RangeInfo(range);
            return rangeInfo;
        }
    }
}
