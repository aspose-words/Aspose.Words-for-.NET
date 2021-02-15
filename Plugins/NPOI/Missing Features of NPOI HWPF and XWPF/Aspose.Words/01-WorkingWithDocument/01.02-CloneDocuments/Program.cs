using System;
using System.IO;
using Aspose.Words;

namespace _01._02_CloneDocuments
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

            // Load a document from the local file system.
            Document doc = new Document("../../data/document.doc");

            // Create another document by cloning the document we loaded.
            Document clone = doc.Clone();

            clone.Save("CloneDocuments.docx");
        }
    }
}
