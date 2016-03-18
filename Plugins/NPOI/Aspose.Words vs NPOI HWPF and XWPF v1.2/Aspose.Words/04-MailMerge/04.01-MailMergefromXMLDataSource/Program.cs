using System;
using System.Collections.Generic;
using System.Text; using Aspose.Words;
using System.IO;
using System.Data;
using System.Reflection;

namespace _04._01_MailMergefromXMLDataSource
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create the Dataset and read the XML.
            DataSet customersDs = new DataSet();
            customersDs.ReadXml("../../data/Customers.xml");

            // Open a template document.
            Document doc = new Document("../../data/TestFile XML.doc");

            // Execute mail merge to fill the template with data from XML using DataTable.
            doc.MailMerge.Execute(customersDs.Tables["Customer"]);

            // Save the output document.
            doc.Save("TestFile XML Out.doc");
        }
    }
}
