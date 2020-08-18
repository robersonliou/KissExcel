using System.Collections.Generic;
using System.Linq;
using KissExcel.Attributes;
using KissExcel.Core;
using KissExcel.Exceptions;
using NSubstitute;
using NUnit.Framework;

namespace KissExcel.Test.Core
{
    [TestFixture]
    public class TestExcelReader
    {
        [SetUp]
        public void Init()
        {
            _excelReader = new ExcelReader();
        }

        [TearDown]
        public void Final()
        {
            _excelReader = null;
        }

        private ExcelReader _excelReader;

        [Test]
        public void ColumnName_Ingore_CaseSensitive()
        {
            var stubExcelReader = new StubExcelReader();
            stubExcelReader.IncludeHeader();
            stubExcelReader.FakeContent = new[] {(0, 0, "Number"), (1, 0, "999")};
            var actual = stubExcelReader.MapTo<IgnoreCaseModel>().First();
            Assert.AreEqual(actual.Id, 999);
        }

        [Test]
        public void ColumnNameAttrMapping_Matched_ColumnName_NotFound_ThrowException()
        {
            var stubExcelReader = new StubExcelReader();
            stubExcelReader.IncludeHeader();
            stubExcelReader.FakeContent = new[] {(0, 0, "Num"), (1, 0, "1")};

            var expectedErrorMessage = "Can not find matched column name:[Number] in the excel header.";

            Assert.Throws(Is.TypeOf<NoMatchedColumnNameException>()
                    .And.Message.EqualTo(expectedErrorMessage),
                () => stubExcelReader.MapTo<ColumnNameAttributeMappingModel>().ToList());
        }

        [Test]
        public void FilePathOfOption_IsNull_ThrowException()
        {
            Assert.Throws<ExcelOptionRequiredException>(() =>
                _excelReader.MapTo<object>());
        }

        [Test]
        public void Mapping_By_Properties()
        {
            var stubExcelReader = new StubExcelReader();
            stubExcelReader.IncludeHeader();
            stubExcelReader.FakeContent = new[] {(0, 0, "Id"), (1, 0, "999")};
            var firstRow = stubExcelReader.MapTo<PropertyMappingModel>().FirstOrDefault();

            Assert.AreEqual(999, firstRow.Id);
        }

        [Test]
        public void Property_Ingore_CaseSensitive()
        {
            var stubExcelReader = new StubExcelReader();
            stubExcelReader.IncludeHeader().IgnoreCase();
            stubExcelReader.FakeContent = new[] {(0, 0, "id"), (1, 0, "999")};
            var actual = stubExcelReader.MapTo<PropertyMappingModel>().First();
            Assert.AreEqual(actual.Id, 999);
        }

        [Test]
        public void PropertyMapping_Matched_ColumnName_NotFound_ThrowException()
        {
            var stubExcelReader = new StubExcelReader();
            stubExcelReader.IncludeHeader();
            stubExcelReader.FakeContent = new[] {(0, 0, "Number"), (1, 0, "1")};

            var expectedErrorMessage = "Can not find matched column name:[Id] in the excel header.";

            Assert.Throws(Is.TypeOf<NoMatchedColumnNameException>()
                    .And.Message.EqualTo(expectedErrorMessage),
                () => stubExcelReader.MapTo<PropertyMappingModel>().ToList());
        }

        [Test]
        public void SheetName_IsNull_ThrowException()
        {
            _excelReader.Open(Arg.Any<string>());
            Assert.Throws<ExcelOptionRequiredException>(() =>
                _excelReader.MapTo<object>());
        }
    }

    internal class StubExcelReader : ExcelReader
    {
        internal IEnumerable<(int rowIndex, int columnIndex, string content)> FakeContent { get; set; }

        protected override void CheckRequiredOptions()
        {
        }

        protected override IEnumerable<(int rowIndex, int columnIndex, string content)> ParseContents()
        {
            RowLength = FakeContent.Count();
            return FakeContent;
        }

        protected override void SetupRequiredMeta()
        {
        }
    }


    internal class PropertyMappingModel
    {
        public int Id { get; set; }
    }

    internal class ColumnNameAttributeMappingModel
    {
        [ColumnName("Number")] public int Id { get; set; }
    }

    internal class IgnoreCaseModel
    {
        [ColumnName("number", IgnoreCase = true)]
        public int Id { get; set; }
    }
}