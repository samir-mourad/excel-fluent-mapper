using ExcelFluentMapper.FluentMapper;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelFluentMapper
{
    public static class ExcelWriter
    {

        public static byte[] ToBytes<TDomain>(this IEnumerable<TDomain> itens, string sheetname, Type type = null)
        {
            var mapClass = type ?? AppDomain.CurrentDomain
                                            .GetAssemblies()
                                            .Where(a => a.FullName.StartsWith("Namespace"))
                                            .SelectMany(i => i.GetTypes().Where(i => i.IsClass && typeof(BaseTypeConfiguration<TDomain>).IsAssignableFrom(i)))
                                            .FirstOrDefault();

            var instance = (BaseTypeConfiguration<TDomain>)Activator.CreateInstance(mapClass);
            var mappings = instance.GetMappings().OrderBy(x => x.Order).ToList();

            var cols = new List<(int index, string name, int order, PropertyConfiguration propertyConfiguration)>();
            var results = new List<object>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add(sheetname);

            var row = 1;

            mappings.ForEach(map => worksheet.Cells[row, map.Order].Value = map.Name);

            foreach (var item in itens)
            {
                row++;

                mappings.ForEach(map =>
                {
                    var propertyInfo = item.GetPropValue(map.Member.Name);
                    var value = Cast(propertyInfo.GetValue(item, null), map);

                    worksheet.Cells[row, map.Order].Value = value;
                });
            }

            return package.GetAsByteArray();

        }

        private static PropertyInfo GetPropValue<TDomain>(this TDomain obj, string propName)
        {
            return obj.GetType().GetProperty(propName);
        }

        private static object Cast(object value, PropertyConfiguration propertyConfiguration)
        {
            if (value == null)
                return value;

            if (propertyConfiguration.WriteFormatFunc != null)
                return propertyConfiguration.WriteFormatFunc(value);

            return value.ToString().Trim();
        }
    }
}
