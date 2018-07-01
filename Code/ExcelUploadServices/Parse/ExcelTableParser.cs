using ExcelUploadServices.InternalInterfaces.Parse;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelUploadServices.Interfaces.Models.Output;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;

namespace ExcelUploadServices.Parse
{
    internal class ExcelTableParser : IExcelTableParser
    {
        private IExcelRowParser excelRowParser;

        public ExcelTableParser(IExcelRowParser excelRowParser)
        {
            this.excelRowParser = excelRowParser;
        }

        public ExcelTable ParseTable(XSSFWorkbook workbook, IName namedRange)
        {
            ExcelTable table = new ExcelTable();
            table.Columns = new List<ExcelColumn>();

            ISheet worksheet = workbook.GetSheet(namedRange.SheetName);

            var startCell = CellRangeAddress.ValueOf(namedRange.RefersToFormula);

            var columnRow = worksheet.GetRow(startCell.FirstRow);

            foreach(var cell in columnRow)
            {
                table.Columns.Add(new ExcelColumn()
                {
                    Name = cell.StringCellValue
                });
            }

            table.Rows = this.excelRowParser.ParseRows(table.Columns, worksheet, startCell);

            return table;
        }
    }
}
