namespace ExcelUploadServices.Interfaces
{
    using System.IO;

    public interface IExcelUploader
    {
        void UploadExcel(byte[] byteArray);
    }
}
