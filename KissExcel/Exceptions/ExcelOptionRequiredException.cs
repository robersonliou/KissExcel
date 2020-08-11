using System;

namespace KissExcel.Exceptions
{
    public class ExcelOptionRequiredException : Exception
    {
        public ExcelOptionRequiredException(string message):base(message)
        {
        }
    }
}