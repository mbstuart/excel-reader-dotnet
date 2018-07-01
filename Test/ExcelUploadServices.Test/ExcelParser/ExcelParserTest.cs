namespace ExcelUploadServices.Test.Parse
{
    using global::ExcelUploadServices.Parse;
    using Interfaces.Models.Output;
    using Interfaces.Parse;
    using InternalInterfaces.Parse;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using Rhino.Mocks;
    using System.IO;

    [TestFixture]
    public class ExcelParserTest
    {
        private ExcelParser excelParser;

        private IExcelTableParser excelTableParser;

        [SetUp]
        protected void SetUp()
        {
            this.excelTableParser = MockRepository.GenerateStub<IExcelTableParser>();

            this.excelTableParser
                .Stub(tableParser => tableParser.ParseTable(Arg<XSSFWorkbook>.Is.Anything, Arg<IName>.Is.Anything))
                .Return(new ExcelTable());

            this.excelParser = new ExcelParser(excelTableParser);
        }

        [Test]
        public void AssertThatExcelParserReturns()
        {
            var excelBytes = TestUtility.RetrieveMockExcelBytes("MockExcel1.xlsx");

            var output = this.excelParser.Parse(excelBytes);

            Assert.IsInstanceOf<ExcelParseOutput>(output);
        }

        [Test]
        public void AssertThatExcelParserReturnsCorrectNumberOfTables()
        {
            var excelBytes = TestUtility.RetrieveMockExcelBytes("MockExcel1.xlsx");

            var output = this.excelParser.Parse(excelBytes);

            Assert.AreEqual(2, output.Tables.Count);
        }

        [Test]
        public void AssertThatExcelParserCallsExcelTableParserEachTime()
        {
            var excelBytes = TestUtility.RetrieveMockExcelBytes("MockExcel1.xlsx");

            var output = this.excelParser.Parse(excelBytes);

            this.excelTableParser
                .AssertWasCalled(tableParser => tableParser.ParseTable(Arg<XSSFWorkbook>.Is.Anything, Arg<IName>.Is.Anything), options => options.Repeat.Times(2));
                
                

        }
    }
}
