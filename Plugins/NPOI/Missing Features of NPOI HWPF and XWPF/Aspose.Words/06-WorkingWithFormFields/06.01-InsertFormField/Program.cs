using Aspose.Words;
using System;
using System.IO;

namespace _06._01_InsertFormField
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

            // Insert a drop down form field with three options.
            string[] items = { "One", "Two", "Three" };
            builder.InsertComboBox("DropDown", items, 0);

            doc.Save("FormFieldTest.docx");
        }
    }
}
