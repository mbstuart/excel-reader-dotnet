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
    public class ExcelTableParser : IExcelTableParser
    {
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

            table.Rows = this.GetRows(table.Columns, worksheet, startCell);

            return table;
        }

        private List<ExcelRow> GetRows(List<ExcelColumn> columns, ISheet worksheet, CellRangeAddress startCell)
        {
            var rows =  new List<ExcelRow>();

            int activeRowNumber = startCell.FirstRow + 1;

            var currentRow = worksheet.GetRow(activeRowNumber);

            while (currentRow != null)
            {
                var row = new ExcelRow();

                foreach(var cell in currentRow)
                {
                    row.Add(cell.StringCellValue);
                }

                rows.Add(row);

                activeRowNumber = activeRowNumber + 1;

                currentRow = worksheet.GetRow(activeRowNumber);
            }

            return rows;
        }
    }
}
