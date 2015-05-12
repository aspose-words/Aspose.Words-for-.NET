//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Data;
using System.IO;
using System.Reflection;

using Aspose.Words;

namespace XMLMailMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Create the Dataset and read the XML.
            DataSet customersDs = new DataSet();
            customersDs.ReadXml(dataDir + "Customers.xml");

            // Open a template document.
            Document doc = new Document(dataDir + "TestFile.doc");

            // Execute mail merge to fill the template with data from XML using DataTable.
            doc.MailMerge.Execute(customersDs.Tables["Customer"]);

            // Save the output document.
            doc.Save(dataDir + "TestFile Out.doc");
        }
    }
}
//ExEnd