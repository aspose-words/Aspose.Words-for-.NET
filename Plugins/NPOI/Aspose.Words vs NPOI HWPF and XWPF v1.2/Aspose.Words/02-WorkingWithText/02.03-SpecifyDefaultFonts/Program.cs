using System;
using System.Collections.Generic;
using System.IO;
using System.Text; using Aspose.Words;
using Aspose.Words.Fonts;

namespace _02._03_SpecifyDefaultFonts
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

            // If the default font defined here cannot be found during rendering then the closest font on the machine is used instead.
            FontSettings.DefaultInstance.DefaultFontName = "Arial Unicode MS";

            // Now the set default font is used in place of any missing fonts during any rendering calls.
            doc.Save("Rendering.SetDefaultFont_Out.pdf");
            doc.Save("Rendering.SetDefaultFont_Out.xps");
        }
    }
}
