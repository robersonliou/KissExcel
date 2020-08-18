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

        public ExcelReader ExcludeHeader()
        {
            _mappingOptions.IncludeHeader = false;
            return this;
        }

        public ExcelReader IgnoreCase()
        {
            _mappingOptions.IgnoreCase = true;
            return this;
        }

        public ExcelReader IncludeHeader()
        {
            _mappingOptions.IncludeHeader = true;
            return this;
        }

        public IEnumerable<TData> MapTo<TData>()
        {
            CheckRequiredOptions();
            SetupRequiredMeta();
            _contentInfos = ParseContents().ToList();

            if (_mappingOptions.IncludeHeader)
                return typeof(TData).PropertiesContainsAttribute<ColumnNameAttribute>()
                    ? MapByHeaderTitles<TData>()
                    : MapByProperties<TData>();

            return Enumerable.Empty<TData>();
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

        private string CellValueLookup(int i, int columnIndex)
        {
            return _contentInfos.SingleOrDefault(x => x.rowIndex == i && x.columnIndex == columnIndex).content;
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

        private IEnumerable<PropertyInfo> GetPropertyInfosWithColumnNameAttr<TData>()
        {
            return typeof(TData).GetProperties().Where(x =>
                x.CustomAttributes.Any(a => a.AttributeType == typeof(ColumnNameAttribute)));
        }

        private (PropertyInfo propertyInfo, string title, int columnIndex) GetPropertyMappingInfo(
            PropertyInfo propertyInfo, string mappingName,
            StringComparison stringComparison = StringComparison.CurrentCulture)
        {
            try
            {
                var headerInfos = _contentInfos.Where(x => x.rowIndex == 0);
                var (_, columnIndex, title) = headerInfos.Single(x => x.content.Equals(mappingName, stringComparison));
                return (propertyInfo, title, columnIndex);
            }
            catch (InvalidOperationException e)
            {
                throw new NoMatchedColumnNameException(
                    $"Can not find matched column name:[{mappingName}] in the excel header.");
            }
        }

        private IEnumerable<(PropertyInfo propertyInfo, string title, int columnIndex)> GetPropertyMappingInfos<TData>()
        {
            var stringComparison = _mappingOptions.IgnoreCase
                ? StringComparison.CurrentCultureIgnoreCase
                : StringComparison.CurrentCulture;
            var propertyMappingInfos = typeof(TData).GetProperties()
                .Select(propertyInfo => GetPropertyMappingInfo(propertyInfo, propertyInfo.Name, stringComparison));
            return propertyMappingInfos;
        }

        private IEnumerable<(PropertyInfo propertyInfo, string columnName, int columnIndex)>
            GetPropertyMappingInfosWithColumnNameAttr<TData>()
        {
            var propertyMappingInfos = GetPropertyInfosWithColumnNameAttr<TData>().Select(propertyInfo =>
            {
                var attribute = propertyInfo.GetCustomAttribute<ColumnNameAttribute>();
                var stringComparison = attribute.IgnoreCase
                    ? StringComparison.CurrentCultureIgnoreCase
                    : StringComparison.CurrentCulture;
                return GetPropertyMappingInfo(propertyInfo, attribute.Name, stringComparison);
            });
            return propertyMappingInfos;
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

        private IEnumerable<TData> MapByHeaderTitles<TData>()
        {
            var propertyMappingInfos = GetPropertyMappingInfosWithColumnNameAttr<TData>().ToList();
            return MapByIndex<TData>(propertyMappingInfos);
        }

        private IEnumerable<TData> MapByIndex<TData>(
            List<(PropertyInfo propertyInfo, string columnName, int columnIndex)> propertyMappingInfos)
        {
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

        private IEnumerable<TData> MapByProperties<TData>()
        {
            var propertyMappingInfos = GetPropertyMappingInfos<TData>().ToList();
            return MapByIndex<TData>(propertyMappingInfos);
        }

        private void ThrowExcelOptionRequiredException(string propertyName)
        {
            throw new ExcelOptionRequiredException(
                $"The {propertyName} property of ExcelMappingOptions is required.");
        }
    }
}