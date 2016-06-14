using Aspose.Words;
using Aspose.Words.MailMerging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class XMLMailMerge
    {
        public static void Run()
        {
            //ExStart:XMLMailMerge 
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 

            // Create the Dataset and read the XML.
            DataSet customersDs = new DataSet();
            customersDs.ReadXml(dataDir + "Customers.xml");

            string fileName = "TestFile XML.doc";
            // Open a template document.
            Document doc = new Document(dataDir + fileName);

            // Execute mail merge to fill the template with data from XML using DataTable.
            doc.MailMerge.Execute(customersDs.Tables["Customer"]);

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the output document.
            doc.Save(dataDir);
            //ExEnd:XMLMailMerge 
            Console.WriteLine("\nMail merge performed with XML data successfully.\nFile saved at " + dataDir);
        }
    }
}
