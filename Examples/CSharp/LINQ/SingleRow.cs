using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace CSharp.LINQ
{
    class SingleRow
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LINQ(); 

            // Load the template document.
            Document doc = new Document(dataDir + "SingleRow.doc");

            // Load the photo and read all bytes.
            byte[] imgdata = System.IO.File.ReadAllBytes(dataDir + "photo.png");
                       
            // Create a Reporting Engine.
            ReportingEngine engine = new ReportingEngine();
            
            // Execute the build report.
            engine.BuildReport(doc, Common.GetManager(), "manager");

            dataDir = dataDir + "SingleRow Out.doc";

            // Save the finished document to disk.
            doc.Save(dataDir);

            Console.WriteLine("\nSingle row template document is populated with the data about manager.\nFile saved at " + dataDir);

        }
    }
}
