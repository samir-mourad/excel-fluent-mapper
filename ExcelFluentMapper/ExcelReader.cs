using ExcelFluentMapper.FluentMapper;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ExcelFluentMapper
{
    public static class ExcelReader
    {
        public static IEnumerable<TDomain> ReadRecords<TDomain>(string path, int headerIndex = 1, string sheet = null, Type type = null)
            where TDomain : new()
        {
            using var package = new ExcelPackage(new FileInfo(path));

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var worksheet = string.IsNullOrWhiteSpace(sheet)
                          ? package.Workbook.Worksheets.FirstOrDefault()
                          : package.Workbook.Worksheets.FirstOrDefault(i => i.Name == sheet);

            var results = worksheet.Read<TDomain>(headerIndex, type);

            return results;
        }
    }

    public static class ExcelReaderExtensions
    {
        public static IEnumerable<TDomain> Read<TDomain>(this ExcelWorksheet sheet, int headerIndex = 1, Type type = null)
            where TDomain : new()
        {
            var results = new List<TDomain>();

            if (sheet == null)
                return results;

            var mapClass = type ?? AppDomain.CurrentDomain
                                            .GetAssemblies()
                                            .Where(a => a.FullName.StartsWith("Namespace"))
                                            .SelectMany(i => i.GetTypes().Where(i => i.IsClass && typeof(BaseTypeConfiguration<TDomain>).IsAssignableFrom(i)))
                                            .FirstOrDefault();

            var instance = (BaseTypeConfiguration<TDomain>)Activator.CreateInstance(mapClass);
            var mappings = instance.GetMappings();

            var rows = sheet.Dimension.Rows;
            var cols = new List<(int index, string name, PropertyConfiguration propertyConfiguration)>();


            for (var i = 1; i < sheet.Dimension.Columns; i++)
            {
                var col = mappings.FirstOrDefault(a => !string.IsNullOrEmpty(a.Name) && a.Name == (string)sheet.Cells[headerIndex, i].Value);

                if (col != null)
                    cols.Add((index: i, name: col.Name, propertyConfiguration: col));
            }

            if (cols.Any())
            {
                mappings.Where(x => x.DefaultValue != null)
                        .ToList()
                        .ForEach(map =>
                        {
                            cols.Add((index: cols.Max(x => x.index) + 1, name: map.Name, propertyConfiguration: map));
                        });
            }

            for (var i = ++headerIndex; i <= rows; i++)
            {
                var t = new TDomain();

                for (var j = 0; j < cols.Count; j++)
                {
                    var property = cols[j];
                    var propertyInfo = t.GetType().GetProperty(property.propertyConfiguration.Member.Name);

                    if (property.propertyConfiguration.DefaultValue != null)
                    {
                        propertyInfo.SetValue(t, property.propertyConfiguration.DefaultValue);
                    }
                    else
                    {
                        var value = Cast(sheet.Cells[i, property.index].Value, propertyInfo.PropertyType, property.propertyConfiguration);

                        if (propertyInfo.PropertyType.IsEnum && value != null)
                            propertyInfo.SetValue(t, Enum.Parse(propertyInfo.PropertyType, value.ToString()));
                        else
                            propertyInfo.SetValue(t, value);
                    }
                }

                results.Add(t);
            }

            return results;
        }

        private static object Cast(object value, Type propertyType, PropertyConfiguration propertyConfiguration)
        {
            if (value == null)
                return value;

            if (propertyType == typeof(byte) || propertyType == typeof(byte?))
            {
                byte.TryParse(propertyConfiguration.StringFormatFunc == null ? value.ToString() : propertyConfiguration.StringFormatFunc(value.ToString()), out byte result);
                return result;
            }
            else if (propertyType == typeof(int))
            {
                int.TryParse(propertyConfiguration.StringFormatFunc == null ? value.ToString() : propertyConfiguration.StringFormatFunc(value.ToString()), out int result);
                return result;
            }
            else if (propertyType == typeof(short))
            {
                short.TryParse(propertyConfiguration.StringFormatFunc == null ? value.ToString() : propertyConfiguration.StringFormatFunc(value.ToString()), out short result);
                return result;
            }
            else if (propertyType == typeof(DateTime) || propertyType == typeof(DateTime?))
            {
                if (propertyConfiguration.DateFormatFunc != null)
                    return propertyConfiguration.DateFormatFunc(value.ToString());

                return value;
            }
            else if (propertyType == typeof(double) || propertyType == typeof(double?)
                  || propertyType == typeof(decimal) || propertyType == typeof(decimal?))
            {
                if (propertyConfiguration.DecimalFormatFunc != null)
                    return propertyConfiguration.DecimalFormatFunc(value.ToString());

                return value.ToString().ToDecimal();

            } else
            {
                if (propertyConfiguration.StringFormatFunc != null)
                    return propertyConfiguration.StringFormatFunc(value.ToString());

                return value.ToString().Trim();
            }
        }

        private static decimal ToDecimal(this string value, CultureInfo cultureInfo = null)
        {
            decimal.TryParse(value, NumberStyles.Any, cultureInfo ?? CultureInfo.InvariantCulture, out var decimalValue);

            return decimalValue;
        }
    }
}
