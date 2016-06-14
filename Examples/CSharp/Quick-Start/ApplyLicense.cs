using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AWords = Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Quick_Start
{
    class ApplyLicense
    {
        public static void Run()
        {

            AWords.License license = new AWords.License();

            // This line attempts to set a license from several locations relative to the executable and Aspose.Words.dll.
            // You can also use the additional overload to load a license from a stream, this is useful for instance when the 
            // license is stored as an embedded resource 
            try
            {
                license.SetLicense("Aspose.Words.lic");
                Console.WriteLine("License set successfully.");
            }

            catch (Exception e)
            {
                // We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license. 
                Console.WriteLine("\nThere was an error setting the license: " + e.Message);
            }
        }
    }
}
