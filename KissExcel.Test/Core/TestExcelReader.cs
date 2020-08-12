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
            stubExcelReader.IncludeHeader(true);
            stubExcelReader.FakeContent = new[] {(0, 0, "Number"), (1, 0, "999")};
            var actual = stubExcelReader.MapTo<IgnoreCaseModel>().First();
            Assert.AreEqual(actual.Id, 999);
        }

        [Test]
        public void FilePathOfOption_IsNull_ThrowException()
        {
            Assert.Throws<ExcelOptionRequiredException>(() =>
                _excelReader.MapTo<object>());
        }

        [Test]
        public void Matched_ColumnName_NotFound_ThrowException()
        {
            var stubExcelReader = new StubExcelReader();
            stubExcelReader.IncludeHeader(true);
            stubExcelReader.FakeContent = new[] {(0, 0, "Num"), (1, 0, "1")};

            var expectedErrorMessage = "Can not find matched column name:[Number] in the excel header.";

            Assert.Throws(Is.TypeOf<NoMatchedColumnNameException>()
                .And.Message.EqualTo(expectedErrorMessage), () => stubExcelReader.MapTo<SimpleModel>().ToList());
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

        protected override void SetupRequiredMeta()
        {
        }

        protected override IEnumerable<(int rowIndex, int columnIndex, string content)> ParseContents()
        {
            RowLength = FakeContent.Count();
            return FakeContent;
        }
    }

    internal class SimpleModel
    {
        [ColumnName("Number")] public int Id { get; set; }
    }

    internal class IgnoreCaseModel
    {
        [ColumnName("number", IgnoreCase = true)]
        public int Id { get; set; }
    }
}