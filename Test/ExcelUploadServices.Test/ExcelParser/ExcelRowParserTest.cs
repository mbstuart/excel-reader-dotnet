

namespace ExcelUploadServices.Test.ExcelParser
{
    using global::ExcelUploadServices.Parse;
    using Interfaces.Models.Output;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    [TestFixture]
    public class ExcelRowParserTest
    {

        private ExcelRowParser excelRowParser;

        private XSSFWorkbook workbook;

        private IName name;

        private List<ExcelColumn> columns;

        [SetUp]
        protected void SetUp()
        {
            this.excelRowParser = new ExcelRowParser();

            var excelBytes = TestUtility.RetrieveMockExcelBytes("MockExcel1.xlsx");
            MemoryStream tableStream = new MemoryStream(excelBytes);

            this.workbook = new XSSFWorkbook(tableStream);

            this.name = workbook.GetName("TableId_Table1");

            this.columns = new List<ExcelColumn>()
            {
                new ExcelColumn()
                {
                    Name = "Column1"
                },
                new ExcelColumn()
                {
                    Name = "Column2"
                },
                new ExcelColumn()
                {
                    Name = "Column3"
                },
            };
        }

        [Test]
        public void ShouldGenerateThreeRowsForExcel()
        {
            var rows = this.excelRowParser.ParseRows(this.columns, this.workbook[0], CellRangeAddress.ValueOf(this.name.RefersToFormula));

            Assert.AreEqual(3, rows.Count);
        }

        [Test]
        public void EachRowShouldHaveSameNumberOfCellsAsColumns()
        {
            var rows = this.excelRowParser.ParseRows(this.columns, this.workbook[0], CellRangeAddress.ValueOf(this.name.RefersToFormula) );

            var columnCount = this.columns.Count;

            foreach (var row in rows)
            {
                Assert.AreEqual(columnCount, row.Count);
            }

        }
    }
}
