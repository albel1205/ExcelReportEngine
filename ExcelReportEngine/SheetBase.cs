using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using ExcelReportEngine.Attributes;
using ExcelReportEngine.Models;
using System.Reflection;
using ExcelReportEngine.Exceptions;
using ExcelReportEngine;

namespace ExcelReportEngine
{
    public abstract class SheetBase
    {
        private Type iLocatableType = typeof(ILocatable);
        private Type iRangeDecoratorType = typeof(IRangeDecorator);
        private Type iSheetDecoratorType = typeof(ISheetDecorator);

        public virtual ExcelWorksheet AddWorksheetTo(ExcelWorkbook workbook)
        {
            var worksheetAttribute =
                Attribute.GetCustomAttribute(this.GetType(), typeof(Worksheet)) as Worksheet;

            var sheetName = worksheetAttribute.Name;
            var worksheet = workbook.Worksheets.Add(sheetName);

            var props = this.GetType().GetProperties();
            foreach (var prop in props)
            {
                var propAttributes = GetPropertyAttributes(prop);

                RangeInfo rangeInfo = GetRangeObject(propAttributes.ToArray(), prop);
                object value = prop.GetValue(this);

                propAttributes.ForEach(x => x.ApplyToSheet(worksheet, rangeInfo, value));
            }

            var sheetAttributes = GetSheetAttributes();
            sheetAttributes.ForEach(x => x.ApplyToSheet(worksheet));

            return worksheet;
        }

        private List<IRangeDecorator> GetPropertyAttributes(PropertyInfo property)
        {
            return property.GetCustomAttributes(true)
                    .Where(x => iRangeDecoratorType.IsInstanceOfType(x))
                    .Cast<IRangeDecorator>()
                    .ToList();
        }

        private List<ISheetDecorator> GetSheetAttributes()
        {
            return this.GetType()
                .GetCustomAttributes()
                .Where(x => iSheetDecoratorType.IsInstanceOfType(x))
                .Cast<ISheetDecorator>()
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
            var value = prop.GetValue(this);
            if (value == null)
            {
                throw new ArgumentNullException(prop.Name);
            }

            var rangeInfo = new RangeInfo(range);
            return rangeInfo;
        }
    }
}
