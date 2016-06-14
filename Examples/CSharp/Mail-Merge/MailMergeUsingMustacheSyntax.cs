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
using System.Web;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class MailMergeUsingMustacheSyntax
    {
        public static void Run()
        {
            //ExStart:MailMergeUsingMustacheSyntax
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 
            DataSet ds = new DataSet();

            ds.ReadXml(dataDir + "Vendors.xml");

            // Open a template document.
            Document doc = new Document(dataDir + "VendorTemplate.doc");

            doc.MailMerge.UseNonMergeFields = true;

            // Execute mail merge to fill the template with data from XML using DataSet.
            doc.MailMerge.ExecuteWithRegions(ds);
            dataDir = dataDir + "MailMergeUsingMustacheSyntax_out_.docx";
            // Save the output document.
            doc.Save(dataDir);
            //ExEnd:MailMergeUsingMustacheSyntax
            Console.WriteLine("\nMail merge performed with mustache syntax successfully.\nFile saved at " + dataDir);
        }
    }
}
