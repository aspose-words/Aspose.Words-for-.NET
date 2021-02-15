using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace _03._03_AutofitSettingtoTables
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

            // Load a document that contains tables.
            Document doc = new Document("../../data/document.doc");

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Autofit the first table to the page width.
            table.AutoFit(AutoFitBehavior.AutoFitToWindow);

            // Save the document to the local file system.
            doc.Save("AutofitSettingtoTables.docx");
        }
    }
}
