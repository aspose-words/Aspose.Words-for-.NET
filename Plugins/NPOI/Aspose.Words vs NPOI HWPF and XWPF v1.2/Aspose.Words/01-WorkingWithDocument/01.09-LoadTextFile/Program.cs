using System;
using System.Collections.Generic;
using System.IO;
using System.Text; using Aspose.Words;

namespace _01._09_LoadTextFile
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


            // The encoding of the text file is automatically detected.
            
            Document doc = new Document("../../data/LoadTxt.txt");

            // Save as any Aspose.Words supported format, such as DOCX.
            doc.Save("AsposeLoadTxt_Out.docx");
        }
    }
}
