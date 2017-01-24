using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CSharp.Quick_Start
{
    class ApplyLicenseFromStream
    {
        public static void Run()
        {
            //ExStart:ApplyLicenseFromStream
            Aspose.Words.License license = new Aspose.Words.License();
            try
            {
                // Initializes a license from a stream 
                MemoryStream stream = new MemoryStream(File.ReadAllBytes(@"Aspose.Words.lic"));
                license.SetLicense(stream);
                Console.WriteLine("License set successfully.");
            }
            catch (Exception e)
            {
                // We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license. 
                Console.WriteLine("\nThere was an error setting the license: " + e.Message);
            }
            //ExEnd:ApplyLicenseFromStream
        }
    }
}
