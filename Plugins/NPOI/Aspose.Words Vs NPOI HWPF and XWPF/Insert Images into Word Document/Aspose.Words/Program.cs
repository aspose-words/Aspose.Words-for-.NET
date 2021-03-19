using System;
using System.IO;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check for license and apply if exists.
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                // Apply Aspose.Words API License.
                Aspose.Words.License license = new Aspose.Words.License();
                // Place license file in Bin/Debug/ Folder.
                license.SetLicense("Aspose.Words.lic");
            }

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");

            builder.InsertImage("../../image/Logo.jpg", 400, 400);
            doc.Save("InsertPicturesInWordAspose.docx");
        }
    }
}
