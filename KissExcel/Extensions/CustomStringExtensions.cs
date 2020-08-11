using System;
using System.Collections.Generic;
using System.Text;

namespace KissExcel.Extensions
{
    internal static class CustomStringExtensions
    {
        public static bool IsNullOrEmpty(this string value)
        {
            return string.IsNullOrEmpty(value);
        }
    }
}
