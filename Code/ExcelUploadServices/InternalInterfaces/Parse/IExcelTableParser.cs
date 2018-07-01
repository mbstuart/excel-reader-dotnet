using ExcelUploadServices.Interfaces.Models.Output;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUploadServices.InternalInterfaces.Parse
{
    public interface IExcelTableParser
    {
        ExcelTable ParseTable(XSSFWorkbook workbook, IName namedRange);
    }
}
