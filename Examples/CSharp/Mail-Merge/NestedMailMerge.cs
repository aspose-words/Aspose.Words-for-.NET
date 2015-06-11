//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace CSharp.Mail_Merge
{
    class NestedMailMerge
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = _RunExamples.GetDataDir_MailMergeAndReporting(); ;
            
            // Create the Dataset and read the XML.
            DataSet pizzaDs = new DataSet();

            // Note: The Datatable.TableNames and the DataSet.Relations are defined implicitly by .NET through ReadXml.
            // To see examples of how to set up relations manually check the corresponding documentation of this sample
            pizzaDs.ReadXml(dataDir + "CustomerData.xml");

            // Open the template document.
            Document doc = new Document(dataDir + "Invoice Template.doc");

            // Execute the nested mail merge with regions
            doc.MailMerge.ExecuteWithRegions(pizzaDs);

            // Save the output to file
            doc.Save(dataDir + "Invoice Out.doc");

            Debug.Assert(doc.MailMerge.GetFieldNames().Length == 0, "There was a problem with mail merge");

            Console.WriteLine("\nMail merge performed with nested data successfully.\nFile saved at " + dataDir + "Invoice Out.doc");
        }
    }
}
