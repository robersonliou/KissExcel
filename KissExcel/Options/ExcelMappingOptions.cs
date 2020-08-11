namespace KissExcel.Options
{
    public class ExcelMappingOptions
    {
        public string FilePath { get; set; }
        public string SheetName { get; set; }
        public bool IncludeHeader { get; set; } = false;
        public bool IsEditable { get; set; } = false;
    }
}