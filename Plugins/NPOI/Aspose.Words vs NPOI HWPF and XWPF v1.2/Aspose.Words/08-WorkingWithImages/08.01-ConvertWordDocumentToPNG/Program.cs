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
            // Check for license and apply if exists
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                // Apply Aspose.Words API License
				Aspose.Words.License license = new Aspose.Words.License();
				// Place license file in Bin/Debug/ Folder
				license.SetLicense("Aspose.Words.lic");
            }

			// define document file location
            string fileDir = "../../data/";

            // load the document.
            Document doc = new Document(fileDir + "document.doc");

            //Create an ImageSaveOptions object to pass to the Save method
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

            // Save each page of the document as Png in data folder
            for (int i = 0; i < doc.PageCount; i++)
            {
                options.PageIndex = i;
                doc.Save(string.Format(i + "WordDocumentAsPNG out.Png", i), options);
            }
        }
    }
}
