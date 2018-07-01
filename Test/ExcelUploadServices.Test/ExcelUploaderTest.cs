namespace ExcelUploadServices.Test
{
    using System;
    using NUnit.Framework;
    using Rhino.Mocks;
    using System.IO;
    using System.Reflection;
    using Interfaces.Parse;

    [TestFixture]
    public class ExcelUploadServices
    {
        private ExcelUploader excelUploader;

        private IExcelParser excelParser;

        [SetUp]
        protected void SetUp()
        {
            this.excelParser = MockRepository.GenerateStub<IExcelParser>();

            this.excelUploader = new ExcelUploader(excelParser);
        }

        [Test]
        public void AssertThatExcelUploaderInitialises ()
        {
            Assert.NotNull(this.excelUploader);
        }

        [Test]
        public void AssertThatUploadFailsIfByteArrayIsNull()
        {
            bool exceptionThrown = false;

            try
            {
                this.excelUploader.UploadExcel(null);
            } 
            catch(ArgumentNullException nullArgumentException)
            {
                exceptionThrown = true;
            }

            Assert.IsTrue(exceptionThrown);
        }


        [Test]
        public void AssertThatExcelParserParseIsCalledWhenUploadExcelIsCalled()
        {
            var excelUpload = TestUtility.RetrieveMockExcelBytes("MockExcel1.xlsx");

            this.excelUploader.UploadExcel(excelUpload);

            this.excelParser.AssertWasCalled(parser => parser.Parse(excelUpload));
        }

    }
}
