using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KissExcel.Extensions
{
    public static class CustomTypeExtensions
    {

        public static bool PropertiesContainsAttribute<TAttribute>(this Type type)
        {
            var propertyInfos = type.GetProperties();
            return propertyInfos.Any(x => x.CustomAttributes.Any(a => a.AttributeType == typeof(TAttribute)));
        }
        public static bool IsNullable(this Type type)
        {
            return Nullable.GetUnderlyingType(type) != null;
        }
    }
}
