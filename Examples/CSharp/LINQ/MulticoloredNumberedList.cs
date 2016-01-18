using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace CSharp.LINQ
{
    class MulticoloredNumberedList
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LINQ();

            // Load the template document.
            Document doc = new Document(dataDir + "MulticoloredNumberedList.doc");
                       
            // Create a Reporting Engine.
            ReportingEngine engine = new ReportingEngine();
            
            // Execute the build report.
            engine.BuildReport(doc, Common.GetClients(), "clients");

            dataDir = dataDir + "MulticoloredNumberedList Out.doc";

            // Save the finished document to disk.
            doc.Save(dataDir);

            Console.WriteLine("\nMulticolored numbered list template document is populated with the data about clients.\nFile saved at " + dataDir);

        }
    }
}
