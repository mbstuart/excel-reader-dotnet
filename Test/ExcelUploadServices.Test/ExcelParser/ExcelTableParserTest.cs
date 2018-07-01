namespace ExcelUploadServices.Test.ExcelParser
{
    using global::ExcelUploadServices.Parse;
    using Interfaces.Models.Output;
    using InternalInterfaces.Parse;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using Rhino.Mocks;
    using System.Collections.Generic;
    using System.IO;

    [TestFixture]
    public class ExcelTableParserTest
    {

        private ExcelTableParser excelTableParser;

        private IExcelRowParser excelRowParser;

        private XSSFWorkbook workbook;

        private IName name;

        [SetUp]
        protected void SetUp()
        {
            this.excelRowParser = MockRepository.GenerateStub<IExcelRowParser>();

            this.excelRowParser
                .Stub(rowParser => rowParser.ParseRows(Arg<List<ExcelColumn>>.Is.Anything, Arg<ISheet>.Is.Anything, Arg<CellRangeAddress>.Is.Anything))
                .Return(new List<ExcelRow>() {
                    new ExcelRow() {
                        
                    },
                    new ExcelRow(),
                    new ExcelRow()
                });

            this.excelTableParser = new ExcelTableParser(this.excelRowParser);

            var excelBytes = TestUtility.RetrieveMockExcelBytes("MockExcel1.xlsx");
            MemoryStream tableStream = new MemoryStream(excelBytes);

            this.workbook = new XSSFWorkbook(tableStream);

            this.name = workbook.GetName("TableId_Table1");
        }

        [Test]
        public void TableMustGetCorrectNumberOfColumns()
        {
            var table = this.excelTableParser.ParseTable(this.workbook, this.name);

            Assert.AreEqual(3, table.Columns.Count);
        }

        [Test]
        public void TableMustGetCorrectNumberOfRows()
        {
            var table = this.excelTableParser.ParseTable(this.workbook, this.name);

            Assert.AreEqual(3, table.Rows.Count);
        }

        [Test]
        public void RowCreationShouldBeDelegatedToRowParserClass()
        {
            var table = this.excelTableParser.ParseTable(this.workbook, this.name);

            this.excelRowParser
               .AssertWasCalled(rowParser => rowParser.ParseRows(Arg<List<ExcelColumn>>.Is.Anything, Arg<ISheet>.Is.Anything, Arg< CellRangeAddress>.Is.Anything));
        }
    }
}
