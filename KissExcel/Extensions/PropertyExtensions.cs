using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace KissExcel.Extensions
{
    public static class PropertyExtensions
    {
        public static bool TryGetAttribute<T>(this PropertyInfo properytInfo, out T attribute) where T : Attribute
        {
            attribute = properytInfo.GetCustomAttribute<T>();
            return properytInfo.GetCustomAttribute<T>() != null;
        }
    }
}
