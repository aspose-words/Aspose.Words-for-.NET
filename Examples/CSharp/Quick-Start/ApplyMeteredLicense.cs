using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.Quick_Start
{
    class ApplyMeteredLicense
    {
        public static void Run()
        {
            try
            {
                //ExStart:ApplyMeteredLicense
                // set metered public and private keys
                Aspose.Words.Metered metered = new Aspose.Words.Metered();
                // Access the setMeteredKey property and pass public and private keys as parameters
                metered.SetMeteredKey("*****", "*****");

                // The path to the documents directory. 
                string dataDir = RunExamples.GetDataDir_QuickStart();

                // Load the document from disk.
                Document doc = new Document(dataDir + "Template.doc");
                //Get the page count of document
                Console.WriteLine(doc.PageCount);
                //ExEnd:ApplyMeteredLicense
            }
            catch (Exception e)
            {
                Console.WriteLine("\nThere was an error setting the license: " + e.Message);
            }
            
        }
    }
}
