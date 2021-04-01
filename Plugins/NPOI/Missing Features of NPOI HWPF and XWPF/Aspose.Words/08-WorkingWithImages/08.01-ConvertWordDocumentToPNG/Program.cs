using System;
using Aspose.Words;
using Aspose.Words.Saving;
using System.IO;

namespace Convert_Doc_to_Png
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

            builder.Writeln("Hello world! This is page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is page 2.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is page 3.");

            // Create an ImageSaveOptions object to pass to the Save method.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

            // Save each page of the document to the local file system as an individual PNG image.
            for (int i = 0; i < doc.PageCount; i++)
            {
                options.PageSet = new PageSet(i);
                doc.Save($"Convert_Doc_to_Png.Page {i}.png", options);
            }
        }
    }
}
