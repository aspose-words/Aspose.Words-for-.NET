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
			builder.Write("Image Before ReSize");
			//insert image from disk
			Shape shape = builder.InsertImage(@"../../data/aspose_Words-for-net.jpg");
			// write text in document
			builder.Write("Image After ReSize ");
			//insert image from disk for resize
			shape = builder.InsertImage(@"../../data/aspose_Words-for-net.jpg");
			// To change the shape size. ( ConvertUtil Provides helper functions to convert between various measurement units. like Converts inches to points.)
			shape.Width = ConvertUtil.InchToPoint(0.5);
			shape.Height = ConvertUtil.InchToPoint(0.5);
			// save new document
            builder.Document.Save("ImageReSize.doc");
        }
    }
}
