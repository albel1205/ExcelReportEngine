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
    [AttributeUsage(AttributeTargets.Property)]
    public class Table : AttributeBase
    {
        private Type iLocatableType = typeof(ILocatable);
        private Type iRangeDecoratorType = typeof(IRangeDecorator);
        private Type iSheetDecoratorType = typeof(ISheetDecorator);

        public override void ApplyToSheet(ExcelWorksheet sheet, RangeInfo range, object value)
        {
            TableBase table = value as TableBase;

            var props = table.GetType().GetProperties();
            foreach (var prop in props)
            {
                if (prop.Name.Contains("TypeId"))//default property of Attribute object
                {
                    continue;
                }

                var propAttributes = GetPropertyAttributes(prop);

                object propValue = prop.GetValue(value);

                propAttributes.ForEach(x => x.ApplyToSheet(sheet, null, propValue));
            }
        }

        private List<IRangeDecorator> GetPropertyAttributes(PropertyInfo property)
        {
            return property.GetCustomAttributes(true)
                    .Where(x => iRangeDecoratorType.IsInstanceOfType(x))
                    .Cast<IRangeDecorator>()
                    .ToList();
        }
    }
}
