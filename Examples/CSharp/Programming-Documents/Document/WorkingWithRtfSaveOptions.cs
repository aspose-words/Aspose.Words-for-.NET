using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class WorkingWithRtfSaveOptions
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            SavingImagesAsWmf(dataDir);
        }

        public static void SavingImagesAsWmf(string dataDir)
        {
            // ExStart:SavingImagesAsWmf 
            string fileName = "TestFile.doc";
            Document doc = new Document(dataDir + fileName);

            RtfSaveOptions saveOpts = new RtfSaveOptions();
            saveOpts.SaveImagesAsWmf = true;

            doc.Save(dataDir + "output.rtf", saveOpts);
            //ExEnd:SavingImagesAsWmf
            Console.WriteLine("\nThe document saved successfully.\nFile saved at " + dataDir);
        }

    }
}
