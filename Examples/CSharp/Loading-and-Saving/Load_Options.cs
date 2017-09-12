using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    class Load_Options
    {
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_QuickStart();

            LoadOptionsUpdateDirtyFields(dataDir);

        }

        public static void LoadOptionsUpdateDirtyFields(string dataDir)
        {
            // ExStart:LoadOptionsUpdateDirtyFields  
            LoadOptions lo = new LoadOptions();

            //Update the fields with the dirty attribute
            lo.UpdateDirtyFields = true;

            //Load the Word document
            Document doc = new Document(dataDir + @"input.docx", lo);

            //Save the document into DOCX
            doc.Save(dataDir + "output.docx", SaveFormat.Docx);
            // ExEnd:LoadOptionsUpdateDirtyFields 
            Console.WriteLine("\nUpdate the fields with the dirty attribute successfully.\nFile saved at " + dataDir);
        }
    }
}
