using KissExcel.Attributes;
using KissExcel.Core;

namespace ConsoleAppCore.Attributes
{
    public class MyIndexColumnParserAttribute : ColumnParserAttribute
    {
        public override string OnParsing(string content)
        {
            return content.Replace("#", string.Empty);
        }

    }

    public class MyIndexParser: IColumnParser
    {
        public string OnParsing(string content)
        {
            return content.Replace("#", string.Empty);
        }
    }
}