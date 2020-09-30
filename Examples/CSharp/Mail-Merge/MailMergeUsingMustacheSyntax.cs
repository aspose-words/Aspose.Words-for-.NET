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
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting();
            MustacheSyntaxUsingDataSet(dataDir);
            UseOfifelseMustacheSyntax(dataDir);
        }

        public static void MustacheSyntaxUsingDataTable(string dataDir)
        {
            //ExStart:MustacheSyntaxUsingDataTable
            // Load a document
            Document doc = new Document(dataDir + @"Test.docx");

            // Loop through each row and fill it with data
            DataTable dataTable = new DataTable("list");
            dataTable.Columns.Add("Number");
            for (int i = 0; i < 10; i++)
            {
                DataRow datarow = dataTable.NewRow();
                dataTable.Rows.Add(datarow);
                datarow[0] = "Number " + i;
            }

            // Activate performing a mail merge operation into additional field types 
            doc.MailMerge.UseNonMergeFields = true;
            doc.MailMerge.ExecuteWithRegions(dataTable);
            doc.Save(dataDir + "MailMerge.Mustache.docx");
            //ExEnd:MustacheSyntaxUsingDataTable
        }

        public static void MustacheSyntaxUsingDataSet(string dataDir)
        {
            // ExStart:MailMergeUsingMustacheSyntax
            DataSet ds = new DataSet();
            ds.ReadXml(dataDir + "Vendors.xml");

            // Open a template document.
            Document doc = new Document(dataDir + "VendorTemplate.doc");

            doc.MailMerge.UseNonMergeFields = true;

            // Execute mail merge to fill the template with data from XML using DataSet.
            doc.MailMerge.ExecuteWithRegions(ds);
            dataDir = dataDir + "MailMergeUsingMustacheSyntax_out.docx";
            // Save the output document.
            doc.Save(dataDir);
            // ExEnd:MailMergeUsingMustacheSyntax
            Console.WriteLine("\nMail merge performed with mustache syntax successfully.\nFile saved at " + dataDir);
        }

        public static void UseOfifelseMustacheSyntax(string dataDir)
        {
            // ExStart:UseOfifelseMustacheSyntax
            // Open a template document.
            Document doc = new Document(dataDir + "UseOfifelseMustacheSyntax.docx");

            doc.MailMerge.UseNonMergeFields = true;

            doc.MailMerge.Execute(new String[] { "GENDER" }, new Object[] { "MALE" });

            dataDir = dataDir + "MailMergeUsingMustacheSyntaxifelse_out.docx";
            // Save the output document.
            doc.Save(dataDir);
            // ExEnd:UseOfifelseMustacheSyntax
            Console.WriteLine("\nMail merge performed with mustache if else syntax successfully.\nFile saved at " + dataDir);
        }
    }
}
