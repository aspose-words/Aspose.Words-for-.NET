using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class SetCompatibilityOptions
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            OptimizeFor(dataDir);
        }

        private static void OptimizeFor(string dataDir)
        {
            string fileName = dataDir + "TestFile.docx";
            // ExStart:OptimizeFor
            Document doc = new Document(fileName);
            doc.CompatibilityOptions.OptimizeFor(Settings.MsWordVersion.Word2016);

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);

            // Save the document to disk.
            doc.Save(dataDir);
            // ExEnd:OptimizeFor      
            Console.WriteLine("\nDocument is optimized for MS Word 2016 successfully.\nFile saved at " + dataDir);
        }
    }
}
