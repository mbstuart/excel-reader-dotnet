namespace ExcelUploadServices.InternalInterfaces.Parse
{
    using ExcelUploadServices.Interfaces.Models.Output;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using System.Collections.Generic;
    internal interface IExcelRowParser
    {
        List<ExcelRow> ParseRows(List<ExcelColumn> columns, ISheet worksheet, CellRangeAddress startCell);
    }
}
