using System;
using Aspose.Words;
using System.IO;

namespace Convert_WordPage_Document_to_MultipageTIFF
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

            string fileDir = "../../data/";
            // open the document 
            Document doc = new Document(fileDir + "document.doc");
            // Save the document as multipage TIFF.
            doc.Save("OutputTiff.tiff");

        }
    }
}
