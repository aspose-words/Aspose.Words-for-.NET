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
                // Place license file in Bin/Debug/ Folder
                license.SetLicense("Aspose.Words.lic");
            }

            Document doc1 = new Document("../../data/doc1.doc");
            Document doc2 = new Document("../../data/doc2.doc");

            Document doc3 = doc1.Clone();
            doc3.AppendDocument(doc2, ImportFormatMode.KeepSourceFormatting);
            doc3.Save("AppendDocuments.docx");
        }
    }
}
