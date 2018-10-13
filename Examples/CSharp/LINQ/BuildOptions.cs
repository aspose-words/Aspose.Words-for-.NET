using Aspose.Words.Reporting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class BuildOptions
    {
        public static void Run()
        {
            RemoveEmptyParagraphs();
        }

        public static void RemoveEmptyParagraphs()
        {
            //ExStart:RemoveEmptyParagraphs
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LINQ();

            string fileName = "template_cleanup.docx";
            // Load the template document.
            Document doc = new Document(dataDir + fileName);

            // Create a Reporting Engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
            engine.BuildReport(doc, Common.GetManagers(), "managers"); 

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the finished document to disk.
            doc.Save(dataDir);
            //ExEnd:RemoveEmptyParagraphs
            Console.WriteLine("\nEmpty paragraphs are removed from the document successfully.\nFile saved at " + dataDir);

        }
    }
}
