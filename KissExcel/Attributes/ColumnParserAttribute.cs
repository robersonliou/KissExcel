using System;
using KissExcel.Core;

namespace KissExcel.Attributes
{
    public abstract class ColumnParserAttribute : Attribute, IColumnParser
    {
        public abstract string OnParsing(string content);
    }
}