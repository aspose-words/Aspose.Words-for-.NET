using Aspose.Words;
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
    class NestedMailMerge
    {
        public static void Run()
        {
            //ExStart:NestedMailMerge
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 
            
            // Create the Dataset and read the XML.
            DataSet pizzaDs = new DataSet();

            // Note: The Datatable.TableNames and the DataSet.Relations are defined implicitly by .NET through ReadXml.
            // To see examples of how to set up relations manually check the corresponding documentation of this sample
            pizzaDs.ReadXml(dataDir + "CustomerData.xml");

            string fileName = "Invoice Template.doc";
            // Open the template document.
            Document doc = new Document(dataDir + fileName);

            // Trim trailing and leading whitespaces mail merge values
            doc.MailMerge.TrimWhitespaces = false;

            // Execute the nested mail merge with regions
            doc.MailMerge.ExecuteWithRegions(pizzaDs);

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the output to file
            doc.Save(dataDir);
            //ExEnd:NestedMailMerge
            Debug.Assert(doc.MailMerge.GetFieldNames().Length == 0, "There was a problem with mail merge");

            Console.WriteLine("\nMail merge performed with nested data successfully.\nFile saved at " + dataDir);
        }
    }
}
