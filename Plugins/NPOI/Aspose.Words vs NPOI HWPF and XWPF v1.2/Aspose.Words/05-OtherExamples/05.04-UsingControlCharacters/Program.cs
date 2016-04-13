using System;
using System.Collections.Generic;
using System.Text; 
using Aspose.Words;
using System.IO;

namespace _05._04_UsingControlCharacters
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

            Document doc = new Document("../../data/document.doc");

            // Enter a dummy field into the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField("MERGEFIELD Field");

            // GetText will retrieve all field codes and special characters
            Console.WriteLine("GetText() Result: " + doc.GetText());

            string text = doc.GetText();
            text = text.Replace(ControlChar.Cr, ControlChar.CrLf);

            Console.WriteLine("Replaced text Result: " + text);
        }
    }
}
