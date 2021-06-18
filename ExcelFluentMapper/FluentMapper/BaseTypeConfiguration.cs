using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace ExcelFluentMapper.FluentMapper
{
    public abstract class BaseTypeConfiguration<T>
    {
        protected IList<PropertyConfiguration> Mappings { get; set; }

        public BaseTypeConfiguration()
        {
            Mappings = new List<PropertyConfiguration>();
        }

        public PropertyConfiguration Map<TProperty>(Expression<Func<T, TProperty>> expression)
        {
            var memberExpression = GetMemberExpression(expression);

            var map = new PropertyConfiguration
            {
                Member = memberExpression.Member
            };

            Mappings.Add(map);

            return map;
        }

        public IList<PropertyConfiguration> GetMappings()
        {
            return Mappings;
        }

        protected internal MemberExpression GetMemberExpression<TProperty>(Expression<Func<T, TProperty>> expression)
        {
            if(!(expression.Body is MemberExpression rootMemberExpression))
            {
                throw new ArgumentException("Expression Inválida.", nameof(expression));
            }

            return rootMemberExpression;
        }
    }
}
