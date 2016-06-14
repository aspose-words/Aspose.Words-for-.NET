using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class BulletedList
    {
        public static void Run()
        {
            //ExStart:BulletedList 
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LINQ();
            string fileName = "BulletedList.doc";
            // Load the template document.
            Document doc = new Document(dataDir + fileName);
            
            // Create a Reporting Engine.
            ReportingEngine engine = new ReportingEngine();
            
            // Execute the build report.
            engine.BuildReport(doc, Common.GetClients(), "clients");

            dataDir = dataDir +  RunExamples.GetOutputFilePath(fileName);

            // Save the finished document to disk.
            doc.Save(dataDir);
            //ExEnd:BulletedList
 
            Console.WriteLine("\nBulleted list template document is populated with the data about clients.\nFile saved at " + dataDir);

        }
    }
}
