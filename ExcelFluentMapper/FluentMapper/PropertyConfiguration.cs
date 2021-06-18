using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace ExcelFluentMapper.FluentMapper
{
    public class PropertyConfiguration
    {
        public MemberInfo Member { get; set; }

        public string Name { get; private set; }

        public int Order { get; set; }

        public object DefaultValue { get; private set; }

        public Func<string, DateTime> DateFormatFunc { get; set; }

        public Func<string, decimal> DecimalFormatFunc { get; set; }

        public Func<string, string> StringFormatFunc { get; set; }

        public Func<object, string> WriteFormatFunc { get; set; }

        public PropertyConfiguration WithColumnName(string name)
        {
            Name = name;
            return this;
        }

        public PropertyConfiguration WithColumnOrder(int order)
        {
            Order = order;
            return this;
        }

        public PropertyConfiguration Format(Func<string, DateTime> func)
        {
            DateFormatFunc += func;
            return this;
        }

        public PropertyConfiguration Format(Func<string, decimal> func)
        {
            DecimalFormatFunc += func;
            return this;
        }

        public PropertyConfiguration Format(Func<string, string> func)
        {
            StringFormatFunc += func;
            return this;
        }

        public PropertyConfiguration FormatWrite(Func<object, string> func)
        {
            WriteFormatFunc += func;
            return this;
        }

        public PropertyConfiguration SetDefaultValue(object value)
        {
            DefaultValue = value;
            return this;
        }

    }
}
