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
        public void FilePathOfOption_IsNull_ThrowException()
        {
            Assert.Throws<ExcelOptionRequiredException>(() =>
                _excelReader.MapTo<object>());
        }

        [Test]
        public void Matched_ColumnName_NotFound_ThrowException()
        {
            var stubExcelReader = new StubExcelReader();
            var expectedErrorMessage = "Can not find matched column name:[Number] in the excel header.";

            Assert.Throws(Is.TypeOf<NoMatchedColumnNameException>()
                .And.Message.EqualTo(expectedErrorMessage), () => stubExcelReader.MapTo<SomeData>().ToList());
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
        protected override void CheckRequiredOptions()
        {
        }

        protected override void SetupRequiredMeta()
        {
        }

        protected override IEnumerable<(int rowIndex, int ColumnIndex, string content)> ParseContents()
        {
            yield return (0, 0, "Num");
            yield return (1, 0, "1");
            yield return (2, 0, "2");
        }
    }

    internal class SomeData
    {
        [ColumnName("Number")] public int Id { get; set; }
    }
}