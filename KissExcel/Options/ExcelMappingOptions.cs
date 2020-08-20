namespace KissExcel.Options
{
    public class ExcelMappingOptions
    {
        /// <summary>
        /// The file path of excel.
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// The sheet name of excel.
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// Represents the excel contains header or not, by default is false.
        /// You should specify it as true except using index mapping.
        /// </summary>
        public bool IncludeHeader { get; set; } = false;

        /// <summary>
        /// It's a global ignore case option for property mapping.
        /// </summary>
        public bool IgnoreCase { get; set; } = false;
        public bool IsEditable { get; set; } = false;
    }
}