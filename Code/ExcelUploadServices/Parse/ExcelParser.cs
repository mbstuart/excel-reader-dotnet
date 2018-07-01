namespace ExcelUploadServices.Parse
{
    using ExcelUploadServices.Interfaces.Parse;
    using Interfaces.Models.Output;
    using InternalInterfaces.Parse;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    using System.Collections.Generic;
    using System.IO;
    using System.Text.RegularExpressions;

    public class ExcelParser : IExcelParser
    {
        private readonly Regex sheetIdentifierRegex = new Regex("TableId_[a-zA-Z0-9]*");

        private IExcelTableParser excelTableParser;

        public ExcelParser(IExcelTableParser excelTableParser)
        {
            this.excelTableParser = excelTableParser;
        }

        public ExcelParseOutput Parse(byte[] excelBytes)
        {
            var output = new ExcelParseOutput();
            output.Tables = new List<ExcelTable>();

            var excel = new XSSFWorkbook(new MemoryStream(excelBytes));
            var names = this.GetNamesFromWorkbook(excel);

            names.ForEach(name =>
            {
                if (this.sheetIdentifierRegex.IsMatch(name.NameName))
                {
                     output.Tables.Add(this.excelTableParser.ParseTable(excel, name));
                }
            });

            return output;
        }

        private List<IName> GetNamesFromWorkbook(XSSFWorkbook workbook)
        {
            List<IName> names = new List<IName>();
            int numberOfNames = workbook.NumberOfNames;
            for (int i = 0; i < numberOfNames; i++)
            {
                names.Add(workbook.GetNameAt(i));
            }

            return names;
        }
    }
}
