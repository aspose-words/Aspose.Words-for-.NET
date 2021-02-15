using System;
using Aspose.Words;
using System.IO;

namespace _05._01_InsertPictureinWordDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check for an Aspose.Words license file in the local file system and apply it, if it exists.
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                Aspose.Words.License license = new Aspose.Words.License();

                // Use the license from the bin/debug/ Folder.
                license.SetLicense("Aspose.Words.lic");
            }

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert an image from a file in the local file system.
            builder.InsertImage("../../data/HumpbackWhale.jpg");

            doc.Save("InsertPictureinWordDocument.docx");
        }
    }
}
