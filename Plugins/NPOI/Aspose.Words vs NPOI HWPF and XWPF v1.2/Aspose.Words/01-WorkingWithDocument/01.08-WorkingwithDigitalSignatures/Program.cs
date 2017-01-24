using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using System.IO;

namespace _01._08_WorkingwithDigitalSignatures
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


            // The path to the document which is to be processed.

            string filePath = "../../data/document.doc";

            FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);
            
            if (info.HasDigitalSignature)
            {
                Console.WriteLine(string.Format("Document {0} has digital signatures, they will be lost if you open/save this document with Aspose.Words.", new FileInfo(filePath).Name));
            }
            else
            {
                Console.WriteLine("Document has no digital signature.");
            }
        }
    }
}
