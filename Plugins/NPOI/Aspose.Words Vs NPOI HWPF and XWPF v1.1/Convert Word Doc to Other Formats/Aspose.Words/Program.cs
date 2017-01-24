using System;
using System.IO;
using Aspose.Words;
namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = Path.GetDirectoryName(Path.GetDirectoryName(Directory.GetCurrentDirectory())) + @"\data\" + "Convert Word Doc to Other Formats.doc";
                     
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
            wordDocument.Save("Convert Word Doc to Other Formatsblank.docx", SaveFormat.Docx);
            wordDocument.Save("Convert Word Doc to Other Formatsblank.bmp", SaveFormat.Bmp);
            wordDocument.Save("Convert Word Doc to Other Formatsblank.html", SaveFormat.Html);
            wordDocument.Save("Convert Word Doc to Other Formatsblank.pdf", SaveFormat.Pdf);
            wordDocument.Save("Convert Word Doc to Other Formatsblank.text", SaveFormat.Text);
        }
    }
}
