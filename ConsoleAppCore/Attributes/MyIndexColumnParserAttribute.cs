using KissExcel.Attributes;

namespace ConsoleAppCore.Attributes
{
    public class MyIndexColumnParserAttribute : ColumnParserAttribute
    {
        public override string OnParsing(string content)
        {
            return content.Replace("#", string.Empty);
        }
    }
}