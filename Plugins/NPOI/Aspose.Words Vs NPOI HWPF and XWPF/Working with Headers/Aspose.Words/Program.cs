using System;
using System.IO;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = Path.GetDirectoryName(Path.GetDirectoryName(Directory.GetCurrentDirectory())) + @"\data\" + "Working with Headers.doc";
            // Check for license and apply if exists
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                // Apply Aspose.Words API License
                Aspose.Words.License license = new Aspose.Words.License();
                // Place license file in Bin/Debug/ Folder
                license.SetLicense("Aspose.Words.lic");
            }


            Document wordDocument = new Document(filePath);
            HeaderFooterCollection headers = wordDocument.FirstSection.HeadersFooters;
            foreach (HeaderFooter header in headers)
            {
                if (header.HeaderFooterType == HeaderFooterType.HeaderFirst || header.HeaderFooterType == HeaderFooterType.HeaderPrimary || header.HeaderFooterType == HeaderFooterType.HeaderEven)
                    Console.WriteLine(header.GetText());
            }
        }
    }
}
