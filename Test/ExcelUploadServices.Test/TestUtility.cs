namespace ExcelUploadServices.Test
{
    using System.IO;
    using System.Reflection;

    internal static class TestUtility
    {
        internal static byte[] RetrieveMockExcelBytes(string fileName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "ExcelUploadServices.Test.ExcelFiles." + fileName;

            byte[] resource = null;
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                MemoryStream ms = new MemoryStream();

                stream.CopyTo(ms);

                resource = ms.ToArray();
            }
            return resource;
        }
    }
}
