using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace CSharp.LINQ
{
    class BubbleChart
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LINQ(); 

            string fileName = "BubbleChart.docx";
            // Load the template document.
            Document doc = new Document(dataDir + fileName);

            // Create a Reporting Engine.
            ReportingEngine engine = new ReportingEngine();
            
            // Execute the build report.
            engine.BuildReport(doc, Common.GetContracts(), "contracts");

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);

            // Save the finished document to disk.
            doc.Save(dataDir);

            Console.WriteLine("\nBubble chart template document is populated with the data about contracts.\nFile saved at " + dataDir);

        }
    }
}
