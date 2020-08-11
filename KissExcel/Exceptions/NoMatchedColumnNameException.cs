using System;

namespace KissExcel.Exceptions
{
    public class NoMatchedColumnNameException : Exception
    {
        public NoMatchedColumnNameException(string message) : base(message)
        {
        }
    }
}