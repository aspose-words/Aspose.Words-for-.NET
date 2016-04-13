using System;
using System.IO;
using Aspose.Words;
namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = Path.GetDirectoryName(Path.GetDirectoryName(Directory.GetCurrentDirectory())) + @"\data\" + "Footers.doc";
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
            HeaderFooterCollection footers = wordDocument.FirstSection.HeadersFooters;
            foreach (HeaderFooter footer in footers)
            {
                if (footer.HeaderFooterType == HeaderFooterType.FooterFirst || footer.HeaderFooterType == HeaderFooterType.FooterPrimary || footer.HeaderFooterType == HeaderFooterType.FooterEven)
                    Console.WriteLine(footer.GetText());
            }

        }
    }
}
