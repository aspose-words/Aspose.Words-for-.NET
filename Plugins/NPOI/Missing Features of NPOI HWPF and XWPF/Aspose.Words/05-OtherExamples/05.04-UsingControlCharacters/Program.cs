using System;
using Aspose.Words;
using System.IO;

namespace _05._04_UsingControlCharacters
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

            Document doc = new Document();

            // Enter a dummy field into the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField("MERGEFIELD Field");

            // "GetText()" will retrieve all field codes and special characters.
            Console.WriteLine("GetText() Result: " + doc.GetText());

            string text = doc.GetText();
            text = text.Replace(ControlChar.Cr, ControlChar.CrLf);

            Console.WriteLine("Replaced text Result: " + text);
        }
    }
}
