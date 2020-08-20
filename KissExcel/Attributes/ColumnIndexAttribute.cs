using System;

namespace KissExcel.Attributes
{
    public class ColumnIndexAttribute : Attribute
    {
        public int Index { get; set; }
        public ColumnIndexAttribute(int index)
        {
            Index = index;
        }
    }
}