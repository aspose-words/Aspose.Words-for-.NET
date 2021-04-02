using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.IO;

namespace Image_ReSize
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

            // Insert an image from the local file system.
            builder.Writeln("Original size:");
            Shape shape = builder.InsertImage(@"../../data/aspose_Words-for-net.jpg");

            builder.InsertParagraph();
            builder.Writeln("Re-sized:");
			shape = builder.InsertImage(@"../../data/aspose_Words-for-net.jpg");

			// To change the shape size.
			// ConvertUtil Provides helper functions to convert between various measurement units, such as inches to points.
			shape.Width = ConvertUtil.InchToPoint(0.5);
			shape.Height = ConvertUtil.InchToPoint(0.5);

            builder.Document.Save("Image_ReSize.docx");
        }
    }
}
