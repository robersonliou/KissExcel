using System;

namespace KissExcel.Attributes
{
    public class ColumnNameAttribute : Attribute
    {
        public string Name { get; set; }

        public bool IgnoreCase { get; set; } = false;

        public ColumnNameAttribute(string name)
        {
            Name = name;
        }
    }
}