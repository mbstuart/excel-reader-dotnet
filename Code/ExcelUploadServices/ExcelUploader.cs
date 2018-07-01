namespace ExcelUploadServices
{
    using System;
    using System.IO;
    using ExcelUploadServices.Interfaces;
    using Interfaces.Parse;

    public class ExcelUploader : IExcelUploader
    {
        private IExcelParser excelParser;

        public ExcelUploader(IExcelParser excelParser)
        {
            this.excelParser = excelParser;
        }

        public void UploadExcel(byte[] byteArray)
        {
            if(byteArray == null)
            {
                throw new ArgumentNullException("memoryStream");
            }

            this.excelParser.Parse(byteArray);

        }
    }
}
