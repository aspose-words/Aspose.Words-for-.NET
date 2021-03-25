using System;
using Aspose.Words;
using System.IO;
 
namespace _01._01_AppendDocuments
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

            // Load two documents from the local file system that we will append together into a new document.
            Document doc1 = new Document("../../data/doc1.doc");
            Document doc2 = new Document("../../data/doc2.doc");

            // Combine the documents by creating a clone of the first document, and then appending the second document to it. 
            Document doc3 = doc1.Clone();
            doc3.AppendDocument(doc2, ImportFormatMode.KeepSourceFormatting);
            doc3.Save("AppendDocuments.docx");
        }
    }
}
