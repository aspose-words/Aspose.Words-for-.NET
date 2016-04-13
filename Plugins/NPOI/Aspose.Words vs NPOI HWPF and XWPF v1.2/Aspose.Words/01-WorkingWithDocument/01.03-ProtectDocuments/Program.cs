using System;
using System.IO;
using Aspose.Words;

namespace _01._03_ProtectDocuments
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
            doc.Protect(ProtectionType.ReadOnly);

            // Following other Protection types are also available
            // ProtectionType.NoProtection
            // ProtectionType.AllowOnlyRevisions
            // ProtectionType.AllowOnlyComments
            // ProtectionType.AllowOnlyFormFields

            doc.Save("AsposeProtect.doc", SaveFormat.Doc);
        }
    }
}
