// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using System.Data;
using System.IO;
using System.Reflection;
using Aspose.Words;

namespace Mail_Merge_from_XML_using_DataSet
{
    class Program
    {
        static void Main(string[] args)
        {
            // Sample infrastructure.
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            string dataDir = new Uri(new Uri(exeDir), @"../../Data/").LocalPath;

            // Create the Dataset and read the XML.
            DataSet customersDs = new DataSet();
            customersDs.ReadXml(dataDir + "Customers.xml");

            // Open a template document.
            Document doc = new Document(dataDir + "Customer Info.doc");

            // Execute mail merge to fill the template with data from XML using DataTable.
            doc.MailMerge.Execute(customersDs.Tables["Customer"]);

            // Save the output document.
            doc.Save(dataDir + "Customer Info.doc");
        }
    }
}
