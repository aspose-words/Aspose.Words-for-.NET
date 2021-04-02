using System;
using Aspose.Words;
using Aspose.Words.Fields;
using System.IO;

namespace _06._02_RemoveFormField
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
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a field that displays the current page number.
            Field field = builder.InsertField("PAGE");

            // Remove the field from the document.
            field.Remove();

            doc.Save("RemoveFormField.docx");
        }
    }
}
