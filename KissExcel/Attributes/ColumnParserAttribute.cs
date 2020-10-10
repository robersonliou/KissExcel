using System;
using System.Net.Mime;
using KissExcel.Core;

namespace KissExcel.Attributes
{
    public class ColumnParserAttribute : Attribute, IColumnParser
    {
        private IColumnParser _columnParser;

        public ColumnParserAttribute() { }

        public ColumnParserAttribute(Type parserType)
        {
            _columnParser = (IColumnParser) Activator.CreateInstance(parserType);
        }

        public virtual string OnParsing(string content)
        {
            return _columnParser.OnParsing(content);
        }
    }
}