using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class InTableMasterDetail
    {
        public static void Run()
        {
            //ExStart:InTableMasterDetail
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LINQ(); 
            string fileName = "InTableMasterDetail.doc";
            // Load the template document.
            Document doc = new Document(dataDir + fileName);

            // Create a Reporting Engine.
            ReportingEngine engine = new ReportingEngine();
            
            // Execute the build report.
            engine.BuildReport(doc, Common.GetManagers(), "managers");

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);

            // Save the finished document to disk.
            doc.Save(dataDir);
            //ExEnd:InTableMasterDetail
            Console.WriteLine("\nIn-Table master detail template document is populated with the data about managers and it's contracts.\nFile saved at " + dataDir);

        }
    }
}
