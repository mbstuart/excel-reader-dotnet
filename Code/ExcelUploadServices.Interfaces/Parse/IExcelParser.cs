namespace ExcelUploadServices.Interfaces.Parse
{
    using ExcelUploadServices.Interfaces.Models.Output;

    public interface IExcelParser
    {
        ExcelParseOutput Parse(byte[] excelBytes);
    }
}
