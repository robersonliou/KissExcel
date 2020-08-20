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
        private List<Row> _rows;
        private SharedStringTable _sharedStringTable;
        private Worksheet _sheet;
        private SpreadsheetDocument _spreadSheetDoc;
        public ExcelMappingOptions MappingOptions { get; set; } = new ExcelMappingOptions();

        public int RowLength { get; set; }
        public int ColumnLength { get; set; }

        public void Dispose()
        {
            _spreadSheetDoc.Dispose();
        }

        public ExcelReader ExcludeHeader()
        {
            MappingOptions.IncludeHeader = false;
            return this;
        }

        public ExcelReader IgnoreCase()
        {
            MappingOptions.IgnoreCase = true;
            return this;
        }

        public ExcelReader IncludeHeader()
        {
            MappingOptions.IncludeHeader = true;
            return this;
        }

        public IEnumerable<TData> MapTo<TData>()
        {
            CheckRequiredOptions();
            SetupRequiredMeta();
            _contentInfos = ParseContents().ToList();

            if (MappingOptions.IncludeHeader)
                return typeof(TData).PropertiesContainsAttribute<ColumnNameAttribute>()
                    ? MapByHeaderTitles<TData>()
                    : MapByProperties<TData>();

            if (typeof(TData).PropertiesContainsAttribute<ColumnIndexAttribute>()) return MapByIndexers<TData>();

            return Enumerable.Empty<TData>();
        }

        public ExcelReader Open(string path)
        {
            MappingOptions.FilePath = path;
            return this;
        }

        public ExcelReader SheetAs(string sheetName)
        {
            MappingOptions.SheetName = sheetName;
            return this;
        }

        protected virtual void CheckRequiredOptions()
        {
            if (MappingOptions.FilePath.IsNullOrEmpty())
                ThrowExcelOptionRequiredException(nameof(ExcelMappingOptions.FilePath));
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
            _spreadSheetDoc = SpreadsheetDocument.Open(MappingOptions.FilePath, MappingOptions.IsEditable);
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

        private IEnumerable<PropertyInfo> GetPropertyInfosWithColumnIndexAttr<TData>()
        {
            var propertyInfos = typeof(TData).GetProperties().Where(x =>
                x.CustomAttributes.Any(a => a.AttributeType == typeof(ColumnIndexAttribute)));
            return propertyInfos;
        }

        private IEnumerable<PropertyInfo> GetPropertyInfosWithColumnNameAttr<TData>()
        {
            return typeof(TData).GetProperties().Where(x =>
                x.CustomAttributes.Any(a => a.AttributeType == typeof(ColumnNameAttribute)));
        }

        private (PropertyInfo propertyInfo, int columnIndex) GetPropertyMappingInfo(
            PropertyInfo propertyInfo, string mappingName,
            StringComparison stringComparison = StringComparison.CurrentCulture)
        {
            try
            {
                var headerInfos = _contentInfos.Where(x => x.rowIndex == 0);
                var (_, columnIndex, title) = headerInfos.Single(x => x.content.Equals(mappingName, stringComparison));
                return (propertyInfo, columnIndex);
            }
            catch (InvalidOperationException e)
            {
                throw new NoMatchedColumnNameException(
                    $"Can not find matched column name:[{mappingName}] in the excel header.");
            }
        }

        private IEnumerable<(PropertyInfo propertyInfo, int columnIndex)> GetPropertyMappingInfos<TData>()
        {
            var stringComparison = MappingOptions.IgnoreCase
                ? StringComparison.CurrentCultureIgnoreCase
                : StringComparison.CurrentCulture;
            var propertyMappingInfos = typeof(TData).GetProperties()
                .Select(propertyInfo => GetPropertyMappingInfo(propertyInfo, propertyInfo.Name, stringComparison));
            return propertyMappingInfos;
        }

        private IEnumerable<(PropertyInfo propertyInfo, int columnIndex)>
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
            return MappingOptions.SheetName.IsNullOrEmpty()
                ? GetFirstSheetId()
                : GetSheetIdByName(MappingOptions.SheetName);
        }

        private StringValue GetSheetIdByName(string sheetName)
        {
            return _spreadSheetDoc.WorkbookPart.Workbook.Descendants<Sheet>()
                .SingleOrDefault(x => x.Name == sheetName)?.Id;
        }

        private IEnumerable<TData> MapByHeaderTitles<TData>()
        {
            var propertyMappingInfos = GetPropertyMappingInfosWithColumnNameAttr<TData>();
            return PropertyIndexMapping<TData>(propertyMappingInfos);
        }

        private IEnumerable<TData> MapByIndexers<TData>()
        {
            var propertyMappingInfos = GetPropertyInfosWithColumnIndexAttr<TData>().Select(propertyInfo =>
            {
                var attribute = propertyInfo.GetCustomAttribute<ColumnIndexAttribute>();
                return (propertyInfo, attribute.Index);
            });
            return PropertyIndexMapping<TData>(propertyMappingInfos);
        }

        private IEnumerable<TData> MapByProperties<TData>()
        {
            var propertyMappingInfos = GetPropertyMappingInfos<TData>();
            return PropertyIndexMapping<TData>(propertyMappingInfos);
        }

        private IEnumerable<TData> PropertyIndexMapping<TData>(
            IEnumerable<(PropertyInfo propertyInfo, int Index)> propertyMappingInfos)
        {
            for (var i = MappingOptions.IncludeHeader ? 1 : 0; i < RowLength; i++)
            {
                var data = Activator.CreateInstance<TData>();
                foreach (var (propertyInfo, columnIndex) in propertyMappingInfos)
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

        private void ThrowExcelOptionRequiredException(string propertyName)
        {
            throw new ExcelOptionRequiredException(
                $"The {propertyName} property of ExcelMappingOptions is required.");
        }
    }
}