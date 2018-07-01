namespace ExcelUploadServices.Test.ExcelParser
{
    using global::ExcelUploadServices.Parse;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using System.IO;

    [TestFixture]
    public class ExcelTableParserTest
    {

        private ExcelTableParser excelTableParser;

        private XSSFWorkbook workbook;

        private IName name;

        [SetUp]
        protected void SetUp()
        {
            this.excelTableParser = new ExcelTableParser();

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
        public void EachRowShouldHaveSameNumberOfCellsAsColumns()
        {
            var table = this.excelTableParser.ParseTable(this.workbook, this.name);

            var columnCount = table.Columns.Count;
            
            foreach(var row in table.Rows)
            {
                Assert.AreEqual(columnCount, row.Count);
            }
        }
    }
}
