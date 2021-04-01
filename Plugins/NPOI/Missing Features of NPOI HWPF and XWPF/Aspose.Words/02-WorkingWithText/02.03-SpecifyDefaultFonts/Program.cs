using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

namespace _02._03_SpecifyDefaultFonts
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

            // If the default font defined here cannot be found during rendering then the closest font on the machine is used instead.
            FontSettings.DefaultInstance.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Courier New";

            DocumentBuilder builder = new DocumentBuilder(doc);

            // Aspose.Words will use that default font in cases where we write text in an inaccessible font, such as this.
            builder.Font.Name = "Non-existent font";
            builder.Write("Hello world!");

            // Now, the set default font is used in place of any missing fonts during any rendering calls.
            doc.Save("SpecifyDefaultFonts.pdf");
            doc.Save("SpecifyDefaultFonts.xps");
        }
    }
}
