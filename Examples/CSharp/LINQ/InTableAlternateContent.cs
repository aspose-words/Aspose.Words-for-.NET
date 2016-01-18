using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace CSharp.LINQ
{
    class InTableAlternateContent
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LINQ(); 

            // Load the template document.
            Document doc = new Document(dataDir + "InTableAlternateContent.doc");
                    
            // Create a Reporting Engine.
            ReportingEngine engine = new ReportingEngine();
            
            // Execute the build report.
            engine.BuildReport(doc, Common.GetContracts(), "contracts");

            dataDir = dataDir + "InTableAlternateContent Out.doc";

            // Save the finished document to disk.
            doc.Save(dataDir);

            Console.WriteLine("\nIn-Table list with alternate content template document is populated with the data about clients and contract price.\nFile saved at " + dataDir);

        }
    }
}
