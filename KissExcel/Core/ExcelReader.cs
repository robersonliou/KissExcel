using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using KissExcel.Attributes;
using KissExcel.Exceptions;
using KissExcel.Extensions;
using KissExcel.Options;

namespace KissExcel.Core
{
    public class ExcelReader : IDisposable
    {
        private List<(int rowIndex, int columnIndex, string content)> _contentInfos;
        private ExcelMappingOptions _mappingOptions;
        private List<Row> _rows;
        private SharedStringTable _sharedStringTable;
        private Worksheet _sheet;
        private SpreadsheetDocument _spreadSheetDoc;


        public ExcelReader()
        {
            _mappingOptions = new ExcelMappingOptions();
        }

        public int RowLength { get; set; }
        public int ColumnLength { get; set; }

        public void Dispose()
        {
            _spreadSheetDoc.Dispose();
        }

        public ExcelReader IncludeHeader()
        {
            _mappingOptions.IncludeHeader = true;
            return this;
        }

        public ExcelReader ExcludeHeader()
        {
            _mappingOptions.IncludeHeader = false;
            return this;
        }

        public IEnumerable<TData> MapTo<TData>()
        {
            CheckRequiredOptions();
            SetupRequiredMeta();
            _contentInfos = ParseContents().ToList();
            if (_mappingOptions.IncludeHeader && typeof(TData).PropertiesContainsAttribute<ColumnNameAttribute>())
                return MapByHeaderName<TData>();

            return null;
        }

        private IEnumerable<TData> MapByHeaderName<TData>()
        {
            var propertyMappingInfos = GetPropertyMappingInfos<TData>().ToList();

            for (var i = 1; i < RowLength; i++)
            {
                var data = Activator.CreateInstance<TData>();
                foreach (var (propertyInfo, _, columnIndex) in propertyMappingInfos)
                {
                    var value = CellValueLookup(i, columnIndex);
                    var convertedType = propertyInfo.PropertyType.IsNullable()
                        ? propertyInfo.PropertyType.GenericTypeArguments[0]
                        : propertyInfo.PropertyType;
                    var convertedValue = Convert.ChangeType(value, convertedType);
                    propertyInfo.SetValue(data, convertedValue);
                }

                yield return data;
            }
        }

        private string CellValueLookup(int i, int columnIndex)
        {
            return _contentInfos.SingleOrDefault(x => x.rowIndex == i && x.columnIndex == columnIndex).content;
        }

        private IEnumerable<(PropertyInfo propertyInfo, string columnName, int columnIndex)>
            GetPropertyMappingInfos<TData>()
        {
            var propertyMappingInfos = GetPropertyInfosWithExcelColumnAttr<TData>().Select(propertyInfo =>
            {
                var attr = propertyInfo.GetCustomAttribute<ColumnNameAttribute>();
                try
                {
                    var stringComparison = attr.IgnoreCase
                        ? StringComparison.CurrentCultureIgnoreCase
                        : StringComparison.CurrentCulture;

                    var (_, columnIndex, content) = _contentInfos.Single(a =>
                        a.rowIndex == 0 && a.content.Equals(attr.Name, stringComparison));
                    return (propertyInfo, content, columnIndex);
                }
                catch (InvalidOperationException e)
                {
                    throw new NoMatchedColumnNameException(
                        $"Can not find matched column name:[{attr.Name}] in the excel header.");
                }
            });
            return propertyMappingInfos;
        }

        private IEnumerable<PropertyInfo> GetPropertyInfosWithExcelColumnAttr<TData>()
        {
            return typeof(TData).GetProperties().Where(x =>
                x.CustomAttributes.Any(a => a.AttributeType == typeof(ColumnNameAttribute)));
        }

        public ExcelReader Open(string path)
        {
            _mappingOptions.FilePath = path;
            return this;
        }

        public void SetOptions(ExcelMappingOptions mappingOptions)
        {
            _mappingOptions = mappingOptions;
        }

        public ExcelReader SheetAs(string sheetName)
        {
            _mappingOptions.SheetName = sheetName;
            return this;
        }

        protected virtual void CheckRequiredOptions()
        {
            if (_mappingOptions.FilePath.IsNullOrEmpty())
                ThrowExcelOptionRequiredException(nameof(ExcelMappingOptions.FilePath));
            if (_mappingOptions.SheetName.IsNullOrEmpty())
                ThrowExcelOptionRequiredException(nameof(ExcelMappingOptions.SheetName));
        }

        private string GetCellText(Cell cell)
        {
            if (cell.ChildElements.Count == 0)
                return null;
            var cellValue = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                cellValue = _sharedStringTable.ChildElements[int.Parse(cellValue)].InnerText;
            return cellValue;
        }

        private StringValue GetFirstSheetId()
        {
            return _spreadSheetDoc.WorkbookPart.Workbook.Descendants<Sheet>().First().Id;
        }

        protected virtual IEnumerable<(int rowIndex, int columnIndex, string content)> ParseContents()
        {
            for (var i = 0; i < RowLength; i++)
            {
                var cells = _rows[i].Descendants<Cell>().ToList();
                for (var j = 0; j < ColumnLength; j++)
                {
                    var content = GetCellText(cells[j]);
                    yield return (i, j, content);
                }
            }
        }

        protected virtual void SetupRequiredMeta()
        {
            _spreadSheetDoc = SpreadsheetDocument.Open(_mappingOptions.FilePath, _mappingOptions.IsEditable);
            _sheet = ((WorksheetPart) _spreadSheetDoc.WorkbookPart.GetPartById(GetSheetId())).Worksheet;
            _sharedStringTable = _spreadSheetDoc.WorkbookPart.SharedStringTablePart.SharedStringTable;
            _rows = _sheet.Descendants<Row>().ToList();
            RowLength = _rows.Count;
            ColumnLength = _rows.First().Descendants<Cell>().Count();
        }

        private StringValue GetSheetId()
        {
            return _mappingOptions.SheetName.IsNullOrEmpty()
                ? GetFirstSheetId()
                : GetSheetIdByName(_mappingOptions.SheetName);
        }

        private StringValue GetSheetIdByName(string sheetName)
        {
            return _spreadSheetDoc.WorkbookPart.Workbook.Descendants<Sheet>()
                .SingleOrDefault(x => x.Name == sheetName)?.Id;
        }

        private void ThrowExcelOptionRequiredException(string propertyName)
        {
            throw new ExcelOptionRequiredException(
                $"The {propertyName} property of ExcelMappingOptions is required.");
        }
    }
}