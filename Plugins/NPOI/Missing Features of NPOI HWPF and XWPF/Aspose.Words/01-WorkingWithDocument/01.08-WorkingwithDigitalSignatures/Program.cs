using System;
using Aspose.Words;
using System.IO;

namespace _01._08_WorkingwithDigitalSignatures
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

            string filePath = "../../data/document.doc";

            // Determine whether this document contains a digital signature.
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);
            
            if (info.HasDigitalSignature)
            {
                Console.WriteLine($"Document {new FileInfo(filePath).Name} has digital signatures, they will be lost if you open/save this document with Aspose.Words.", new FileInfo(filePath).Name);
            }
            else
            {
                Console.WriteLine("Document has no digital signature.");
            }
        }
    }
}
