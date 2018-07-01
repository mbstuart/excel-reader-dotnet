using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelUploadServices.Interfaces.Models.Output;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using ExcelUploadServices.InternalInterfaces.Parse;

namespace ExcelUploadServices.Parse
{
    internal class ExcelRowParser: IExcelRowParser
    {
        public List<ExcelRow> ParseRows(List<ExcelColumn> columns, ISheet worksheet, CellRangeAddress startCell)
        {
            var rows = new List<ExcelRow>();

            int activeRowNumber = startCell.FirstRow + 1;

            var currentRow = worksheet.GetRow(activeRowNumber);

            while (currentRow != null)
            {
                var row = new ExcelRow();

                foreach (var cell in currentRow)
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
