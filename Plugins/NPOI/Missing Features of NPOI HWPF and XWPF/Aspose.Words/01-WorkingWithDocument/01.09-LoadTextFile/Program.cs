using System;
using System.IO;
using Aspose.Words;

namespace _01._09_LoadTextFile
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

            // Load a plaintext file. Aspose.Words will determine its encoding automatically.
            Document doc = new Document("../../data/LoadTxt.txt");

            // Save the document to the DOCX format.
            doc.Save("AsposeLoadTxt_Out.docx");
        }
    }
}
