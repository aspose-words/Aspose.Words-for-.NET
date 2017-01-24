using System;
using System.Drawing;
using System.IO;
using Aspose.Words;

namespace Insert_Hyperlink_in_document
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check for license and apply if exists
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                // Apply Aspose.Words API License
                Aspose.Words.License license = new Aspose.Words.License();
                // Place license file in Bin/Debug/ Folder
                license.SetLicense("Aspose.Words.lic");
            }


			Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Please make sure to visit ");

            // Specify font formatting for the hyperlink.
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;
            // Insert the link.
            builder.InsertHyperlink("Aspose Website", "http://www.aspose.com", false);

            // Revert to default formatting.
            builder.Font.ClearFormatting();

            builder.Write(" for more information.");
            doc.Save("Insert_Hyperlink_In_Document.doc");

        }
    }
}
