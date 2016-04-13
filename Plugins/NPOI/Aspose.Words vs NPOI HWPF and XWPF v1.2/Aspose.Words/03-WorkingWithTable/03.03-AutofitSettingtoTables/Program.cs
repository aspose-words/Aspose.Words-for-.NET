using System;
using System.Collections.Generic;
using System.IO;
using System.Text; using Aspose.Words;
using Aspose.Words.Tables;

namespace _03._03_AutofitSettingtoTables
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

            // Open the document
            Document doc = new Document("../../data/document.doc");

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Autofit the first table to the page width.
            table.AutoFit(AutoFitBehavior.AutoFitToWindow);

            // Save the document to disk.
            doc.Save("TestFile.AutoFitToWindow Out.doc");
        }
    }
}
