using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class InTableWithFilteringGroupingSorting
    {
        public static void Run()
        {
            //ExStart:InTableWithFilteringGroupingSorting
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LINQ(); 
            string fileName = "InTableWithFilteringGroupingSorting.doc";
            // Load the template document.
            Document doc = new Document(dataDir + fileName);

            // Create a Reporting Engine.
            ReportingEngine engine = new ReportingEngine();
            
            // Execute the build report.
            engine.BuildReport(doc, Common.GetContracts(), "contracts");

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);

            // Save the finished document to disk.
            doc.Save(dataDir);
            //ExEnd:InTableWithFilteringGroupingSorting

            Console.WriteLine("\nIn-Table filtering, grouping and sorting template document is populated with the data about managers.\nFile saved at " + dataDir);

        }
    }
}
