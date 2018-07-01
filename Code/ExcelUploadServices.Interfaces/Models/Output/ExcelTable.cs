
namespace ExcelUploadServices.Interfaces.Models.Output
{
    using System;
    using System.Collections.Generic;
    public class ExcelTable
    {
        public string Id { get; set; }

        public DateTime UploadDate { get; set; }

        public DateTime AsOfDate { get; set; }

        public List<ExcelColumn> Columns { get; set; }

        public List<ExcelRow> Rows { get; set; }
    }
}